import json
import os
from typing import Optional, Any

import requests
from requests.adapters import HTTPAdapter
from requests.cookies import RequestsCookieJar
from urllib3 import Retry

import config
from util import app

global_requests_session = None


class NoCookieJar(requests.cookies.RequestsCookieJar):
    def set_cookie(self, *args, **kwargs):
        pass

    def update(self, *args, **kwargs):
        pass


def create_session_with_retry() -> requests.Session:
    """创建一个带有内置重试机制的 requests.Session 对象。"""
    session = requests.Session()
    # session.trust_env = False

    retry_strategy = Retry(
        # 总重试次数
        total=config.get_config().get("retries", 0),
        # 重试之间的等待时间
        backoff_factor=config.get_config().get("retryDelay", 0),
        # 需要重试的HTTP状态码
        status_forcelist=[408, 429, 500, 502, 503, 504],
        # 允许重试的请求方法
        allowed_methods=["HEAD", "GET", "POST", "OPTIONS", "PUT", "DELETE", "PATCH", "CONNECT"],
    )
    adapter = HTTPAdapter(max_retries=retry_strategy)

    session.mount("https://", adapter)
    session.mount("http://", adapter)

    # 禁止自动管理 Cookie
    session.cookies = NoCookieJar()

    # port = os.getenv("PROXY_PORT", '').strip()
    #
    # proxies = {
    #     "http": f"http://127.0.0.1:{port}",  # 10808, 7897
    #     "https": f"http://127.0.0.1:{port}",
    # }
    # session.proxies = proxies if len(port) > 0 and int(port) > 0 else {}

    return session


def get_requests_session():
    global global_requests_session

    if global_requests_session is None:
        global_requests_session = create_session_with_retry()

    return global_requests_session


def get_proxy_port():
    port = os.getenv("PROXY_PORT", '').strip()

    proxies = {
        "http": f"http://127.0.0.1:{port}",  # 10808, 7897
        "https": f"http://127.0.0.1:{port}",
    }

    return proxies if len(port) > 0 and int(port) > 0 else {}


def get(
        url,
        params=None,
        headers=None,
        timeout=None,
        stream=None
):
    response = requests.get(
        url,
        params=params,
        headers=headers,
        timeout=timeout if timeout else config.get_config().get("timeout", 60),
        stream=stream,
        proxies=get_proxy_port(),
    )
    response.raise_for_status()

    return response


def post(
        url,
        json_data=None,
        headers=None,
        timeout=None,
        stream=None
):
    response = requests.post(
        url,
        json=json_data,
        headers=headers,
        timeout=timeout if timeout else config.get_config().get("timeout", 60),
        stream=stream,
        proxies=get_proxy_port(),
    )
    response.raise_for_status()

    return response


def get_update_info() -> tuple[Optional[dict[str, Any]], Optional[Exception]]:
    try:
        res = get("https://cnb.cool/feassh/zying-auto/-/git/raw/main/version.json", timeout=60)
        return res.json(), None
    except Exception as e:
        return None, e


def check_need_update() -> Optional[tuple[dict[str, Any], str]]:
    current_version = app.get_version()
    if len(current_version) == 0:
        return None

    info, _ = get_update_info()
    if not info:
        return None

    latest_version = info.get("version")

    if latest_version == current_version:
        return None

    return info, current_version


def save_kw_to_server(kws) -> Optional[Exception]:
    if kws is None or len(kws) == 0:
        return None

    data = []
    data_product = []
    task_id = int(config.get_config()["lastTaskId"])

    for kw, kw_img, filter_criteria, products in kws:
        data.append({
            "kw": kw,
            "img": kw_img,
            "filter_criteria": filter_criteria,
            "task_id": task_id
        })

        for (cate_main, cate_sub, fulfiller_type), (asin, img, title, price_symbol, price, buy_number) in products:
            data_product.append({
                "asin": asin,
                "kw": kw,
                "price_symbol": price_symbol,
                "price": price,
                "buy_number": buy_number,
                "delivery": fulfiller_type,
                "title": title,
                "img": img,
                "category_main": cate_main,
                "category_sub": cate_sub,
                "task_id": task_id
            })

    try:
        resp = post("https://zying.feassh.workers.dev/insertBatch", json_data={
            "data": data,
            "productData": data_product,
            "token": "feassh-zying-cf-worker-token"
        })

        data = resp.json()
        if data.get("code", -1) == 0:
            return None
        else:
            return Exception(json.dumps(data))
    except Exception as e:
        return e


def delete_all_server_data() -> Optional[Exception]:
    try:
        resp = post("https://zying.feassh.workers.dev/deleteAll", json_data={
            "token": "feassh-zying-cf-worker-token"
        })

        data = resp.json()
        if data.get("code", -1) == 0:
            return None
        else:
            return Exception(json.dumps(data))
    except Exception as e:
        return e


def get_amz123_kw_list(page) -> tuple[Optional[tuple[list, int]], Optional[Exception]]:
    try:
        res = post(
            "https://api.amz123.com/search/v1/hotwords/search",
            json_data={
                "word": "",
                "country": "jp",
                "ranking_this_week": [50001],
                "fluctuation_range": [1001],
                "word_len_range": [],
                "click_range": [],
                "conversion_range": [],
                "ne_word": "",
                "top3_brand": "",
                "top3_category": "",
                "fluctuation_use_abs": 1,
                "page": {
                    "size": 200,
                    "num": page,
                    "sorts": [{"condition": "new_rank", "order": 1}]
                }
            }
        )

        data = res.json()
        if data.get("status", -1) == 0:
            data2 = data.get("data")
            if not data2:
                return None, Exception(json.dumps(data))
            else:
                return (data2.get("rows", []), data2.get("total", 0)), None
        else:
            return None, Exception(json.dumps(data))
    except Exception as e:
        return None, e
