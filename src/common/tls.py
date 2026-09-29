"""Выключенная проверка TLS объявляется ОДИН раз, а не на каждый запрос.

Оба клиента ходят через нестабильный туннель и сознательно передают
`verify=False`; urllib3 на каждый такой запрос печатает InsecureRequestWarning,
и в логе прогона (который у сдачи запускает эксперт) двадцать одинаковых
абзацев предупреждения прячут настоящие сообщения. Глушить предупреждение
молча нельзя — тогда отключённая проверка нигде не видна. Поэтому: одна
строка на хост при первом обращении и тишина дальше.
"""

import sys
from urllib.parse import urlparse

import urllib3

_announced = set()


def allow_unverified_tls(url_or_host, stream=None):
    host = urlparse(url_or_host).hostname or str(url_or_host)
    if host in _announced:
        return
    _announced.add(host)
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    print(f"! TLS-проверка отключена для {host} (туннель без валидного сертификата)",
          file=stream or sys.stderr, flush=True)
