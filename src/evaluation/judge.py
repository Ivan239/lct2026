"""LLM half of the rubric scorer — the criteria a machine genuinely can't call:
composition balance, colour harmony, relevance to the brief, hallucinations.

Two calls per deck, split by what each criterion actually needs:
  - a vision call (whole deck attached at once — GigaChat-2-Max reads a multi-
    image message fine) for everything you have to SEE;
  - a text call (brief + extracted slide text, no images, cheaper text tier) for
    everything you have to READ.

The model returns strict JSON; shape is enforced in code (ids validated, scores
clamped to 1..5 or N/A) per the project rule that the model doesn't follow format
instructions on trust. A hard judge failure degrades that group to N/A with a
recorded note rather than crashing — an autonomous loop must survive a flaky
judge, and the deterministic half still produces a real partial score.
"""

import time

import requests

from common.json_utils import extract_json
from common.model_fallback import (
    NETWORK_RETRY_DELAY_SECONDS,
    TEXT_MODELS,
    TRANSIENT_NETWORK_ERRORS,
    VISION_MODELS,
    call_with_model_fallback,
)
from evaluation.rubric import CRITERIA

# Ask for at most this many criteria per call. GigaChat-2-Max answers a small,
# focused criteria set as clean JSON every time; asking for all 22 visual
# criteria in one shot came back missing keys (verified against a real 7-slide
# deck). Chunking trades a couple extra cheap calls for reliable output.
CHUNK = 7

# One-line definition per criterion to steer the 1..5 judgement. Kept short so
# the whole rubric fits the prompt without drowning the images.
DEFS = {
    "1.2": "изображения не закрывают важный текст (частичное декоративное наложение — ок)",
    "1.3": "баланс композиции: нет перегруза, элементы распределены, взгляд движется естественно",
    "1.4": "свободное пространство: нет огромных пустот и нет чрезмерной плотности",
    "1.5": "выравнивание: объекты по сетке, одинаковые отступы, ничего не «прыгает»",
    "1.6": "визуальная иерархия: понятно где заголовок, ключевая мысль, детали",
    "1.7": "единый стиль: одинаковые шрифты/формы/палитра/иконки по всей деке",
    "1.8": "цветовая гармония: цвета сочетаются, нет случайных ярких пятен, акценты уместны",
    "2.2": "информативность заголовка: после прочтения понятно, о чём слайд",
    "3.2": "структура текста: абзацы/списки/таблицы там, где нужно",
    "3.3": "ясность: нет громоздких конструкций",
    "3.4": "отсутствие повторов одинаковых мыслей",
    "3.5": "грамматика: нет опечаток, орфографических и пунктуационных ошибок",
    "4.1": "изображения соответствуют теме и помогают понять материал",
    "4.2": "качество изображений: нет пикселизации, артефактов ИИ, обрезанных объектов",
    "4.3": "соотношение размеров изображений уместное (не гигантские, не микроскопические)",
    "4.4": "если изображение требует пояснения — подпись есть",
    "4.5": "нет бессмысленных декоративных картинок",
    "5.1": "слайды отвечают поставленной в брифе задаче",
    "5.2": "полнота: ключевые темы брифа не пропущены",
    "5.3": "логика: проблема → анализ → решение → вывод (по смыслу)",
    "5.4": "фактическая корректность данных",
    "5.5": "одни и те же понятия называются одинаково во всей деке",
    "5.6": "нет выдуманных цитат, исследований, статистики, источников",
    "6.1": "логичный порядок слайдов, нет резких смысловых скачков",
    "6.2": "связность: каждый слайд продолжает предыдущий",
    "7.1": "диаграммы корректно отражают данные (N/A если диаграмм нет)",
    "7.2": "понятность инфографики без пояснений (N/A если нет)",
    "7.3": "подписи инфографики читаемы и не перекрываются (N/A если нет)",
    "8.1": "материал соответствует уровню аудитории из брифа",
    "8.2": "соответствие стилю (деловой/академический/маркетинговый/…) из брифа",
    "9.2": "нет наложений элементов: иконки не перекрывают текст/таблицы",
    "9.3": "консистентность: одинаковые поля, размеры заголовков, стили по всем слайдам",
    "10.1": "готовность без ручной правки (5 — можно показывать сразу; 1 — проще заново)",
    "dop_prompt": "точное следование пользовательскому промпту/брифу",
    "dop_brand": "соблюдение фирменного стиля: бренд-цвета и шрифты выдержаны",
    "dop_template": "корректное использование шаблона презентации (не сломан макет)",
    "dop_emphasis": "уместность выделений (жирный/цвет/акценты не хаотичны)",
    "dop_icons": "качественные иконки вместо случайного клипарта",
    "dop_image_style": "согласованность стиля изображений (все фото или все иллюстрации)",
}

# Split by what the criterion needs to be judged.
VISUAL_IDS = [
    "1.2", "1.3", "1.4", "1.5", "1.6", "1.7", "1.8", "9.2", "9.3",
    "dop_emphasis", "dop_icons", "dop_image_style", "dop_brand", "dop_template",
    "4.1", "4.2", "4.3", "4.4", "4.5", "7.1", "7.2", "7.3",
]
CONTENT_IDS = [
    "2.2", "3.2", "3.3", "3.4", "3.5",
    "5.1", "5.2", "5.3", "5.4", "5.5", "5.6",
    "6.1", "6.2", "8.1", "8.2", "10.1", "dop_prompt",
]

_llm = {c for c, (_, mode) in CRITERIA.items() if mode == "llm"}
assert set(VISUAL_IDS) | set(CONTENT_IDS) == _llm, "judge split must cover exactly the llm criteria"
assert not (set(VISUAL_IDS) & set(CONTENT_IDS)), "a criterion can't be in both judge groups"

_SCALE = (
    "Шкала для КАЖДОГО критерия: 5 отлично, 4 хорошо (мелкие недочёты), "
    "3 удовлетворительно (заметно, но рабоче), 2 плохо, 1 очень плохо. "
    'Если критерий неприменим к этой деке — верни строку "NA". '
)


def _criteria_block(ids):
    return "\n".join(f'  "{cid}": {DEFS[cid]}' for cid in ids)


def _out_spec(ids):
    return (
        'Ответь СТРОГО JSON без пояснений и без markdown, формат: '
        '{"<id>": {"score": <1..5 или "NA">, "note": "<кратко, до 12 слов>"}, ...}. '
        f'Оцени ровно эти критерии: {", ".join(ids)}.'
    )


def _require_ids(parsed, ids):
    """GigaChat sometimes returns valid JSON that silently omits a few of the
    requested criteria. Treat a missing key like malformed output — raise so
    call_with_model_fallback retries the same model for a COMPLETE answer,
    rather than accepting a dict with holes."""
    if not isinstance(parsed, dict):
        raise ValueError("judge response is not a JSON object")
    missing = [c for c in ids if c not in parsed]
    if missing:
        raise ValueError(f"judge omitted criteria: {missing}")
    return parsed


def _coerce(raw, ids):
    """Validate the model's dict into {id: {"score": 1..5|None, "note": str}}.
    Missing/garbage entries become N/A with a note — never a fake number."""
    out = {}
    for cid in ids:
        entry = raw.get(cid) if isinstance(raw, dict) else None
        score, note = None, "не оценено моделью"
        if isinstance(entry, dict):
            s = entry.get("score")
            note = str(entry.get("note", ""))[:160]
            if isinstance(s, (int, float)) and not isinstance(s, bool):
                score = max(1, min(5, int(round(s))))
            elif isinstance(s, str) and s.strip().isdigit():
                score = max(1, min(5, int(s.strip())))
            # anything else (incl. "NA") stays None -> N/A
        out[cid] = {"score": score, "detail": note}
    return out


def _chunks(ids, size=CHUNK):
    for i in range(0, len(ids), size):
        yield ids[i:i + size]


JUDGE_UNAVAILABLE = "судья недоступен"  # prefix marking an N/A caused by a judge FAILURE (not inapplicability)


def _judge_group(build_call, ids, models):
    """Score `ids` in reliable-sized chunks. A failing chunk degrades only its
    own ids to N/A — the rest of the deck still gets a real score."""
    out = {}
    for chunk in _chunks(ids):
        try:
            raw = call_with_model_fallback(build_call(chunk), models)
            out.update(_coerce(raw, chunk))
        except Exception as e:  # noqa: BLE001 — a flaky judge must not kill the loop
            out.update({cid: {"score": None, "detail": f"{JUDGE_UNAVAILABLE}: {type(e).__name__}"}
                        for cid in chunk})
    return out


def _unavailable_ids(scores, ids):
    """Of `ids`, those left N/A specifically by a judge FAILURE (transient network
    drop), not by inapplicability — the ones worth a second attempt."""
    return [c for c in ids
            if scores.get(c, {}).get("score") is None
            and str(scores.get(c, {}).get("detail", "")).startswith(JUDGE_UNAVAILABLE)]


def judge_visual(client, image_ids, brief, models=VISION_MODELS, ids=None, context=""):
    ids = VISUAL_IDS if ids is None else ids

    def build(chunk):
        prompt = (
            f"Ты строгий дизайн-ревьюер презентаций. К сообщению прикреплены {len(image_ids)} "
            "слайдов сгенерированной презентации по порядку.\n"
            f"Бриф, по которому её делали:\n{brief}\n\n"
            f"{context}{_SCALE}\nОцени визуальные критерии (смотри на слайды):\n"
            f"{_criteria_block(chunk)}\n\n{_out_spec(chunk)}"
        )

        def call(model):
            r = client.chat(
                [{"role": "user", "content": prompt, "attachments": image_ids}],
                model=model, temperature=0.1, max_tokens=1500,
            )
            return _require_ids(extract_json(r["choices"][0]["message"]["content"]), chunk)

        return call

    return _judge_group(build, ids, models)


def judge_content(client, brief, slide_texts, models=TEXT_MODELS, ids=None):
    ids = CONTENT_IDS if ids is None else ids
    deck = "\n\n".join(f"[Слайд {i + 1}]\n{t}" for i, t in enumerate(slide_texts))

    def build(chunk):
        prompt = (
            "Ты строгий редактор-ревьюер презентаций. Ниже бриф и извлечённый текст слайдов.\n"
            f"Бриф:\n{brief}\n\nТекст слайдов:\n{deck}\n\n"
            f"{_SCALE}\nОцени содержательные критерии (по тексту и брифу):\n"
            f"{_criteria_block(chunk)}\n\n{_out_spec(chunk)}"
        )

        def call(model):
            r = client.chat([{"role": "user", "content": prompt}], model=model,
                            temperature=0.1, max_tokens=1500)
            return _require_ids(extract_json(r["choices"][0]["message"]["content"]), chunk)

        return call

    return _judge_group(build, ids, models)


def _upload_with_retry(client, path, attempts=4):
    """The /files endpoint intermittently 403s / drops on the flaky Sber tunnel
    (verified: same upload succeeds on immediate retry). Treat 403/429/5xx and
    network faults as transient here."""
    last = None
    for i in range(attempts):
        try:
            return client.upload_file(path)["id"]
        except requests.HTTPError as e:
            last = e
            code = e.response.status_code if e.response is not None else None
            if code not in (403, 429, 500, 502, 503, 504):
                raise
        except TRANSIENT_NETWORK_ERRORS as e:
            last = e
        time.sleep(NETWORK_RETRY_DELAY_SECONDS * (i + 1))
    raise last


def judge(client, image_paths, brief, slide_texts, skip_ids=None, media=None,
          vision_models=VISION_MODELS, text_models=TEXT_MODELS):
    """Full LLM pass. Uploads images once, runs both groups, returns
    {criterion_id: {"score": 1..5|None, "detail": str}} for every llm criterion.

    skip_ids: criteria the deck doesn't exercise (e.g. image criteria on a
    text-only deck). They're forced to N/A and never sent to the model — the
    rubric says N/A is 'неприменимо', and the judge scores 'no images' as a 1
    too often to be trusted with that call.
    media: deck_media() result — used to tell the vision judge when 'images' are
    skeleton PLACEHOLDERS (dashed frames with a caption), so it scores their
    PLACEMENT/size/relevance, not a photo that isn't there yet."""
    skip = set(skip_ids or [])
    context = ""
    if media and media.get("placeholders", 0) > 0 and media.get("substantive_pictures", 0) == 0:
        context = (
            "ВАЖНО: изображения на слайдах — это СКЕЛЕТЫ-ЗАГЛУШКИ (пунктирная рамка с "
            "подписью, что там будет), картинки ещё не сгенерированы. Оценивай РАСПОЛОЖЕНИЕ, "
            "размер и уместность этих областей в композиции, а не качество картинки.\n"
        )
    image_ids = [_upload_with_retry(client, p) for p in image_paths]
    vis_ids = [i for i in VISUAL_IDS if i not in skip]
    con_ids = [i for i in CONTENT_IDS if i not in skip]
    scores = {}
    scores.update(judge_visual(client, image_ids, brief, models=vision_models, ids=vis_ids, context=context))
    scores.update(judge_content(client, brief, slide_texts, models=text_models, ids=con_ids))

    # A transient tunnel drop during ONE chunk N/A'd its whole chunk — a real run
    # lost 6 criteria (a full visual chunk) to a momentary ConnectionError. By now
    # the other chunks are done and the blip has almost certainly passed, so give
    # the judge-FAILED criteria (not the inapplicable ones) a single second pass.
    retry_vis = _unavailable_ids(scores, vis_ids)
    retry_con = _unavailable_ids(scores, con_ids)
    if retry_vis:
        scores.update(judge_visual(client, image_ids, brief, models=vision_models, ids=retry_vis, context=context))
    if retry_con:
        scores.update(judge_content(client, brief, slide_texts, models=text_models, ids=retry_con))

    for cid in skip:
        scores[cid] = {"score": None, "detail": "неприменимо: в деке нет такого медиа"}
    return scores
