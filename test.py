import os
import re
from datetime import date

import requests
from bs4 import BeautifulSoup

import config
import util


# def a():
#     try:
#         print("正在尝试获取最新的亚马逊 Cookie...")
#
#         # 先获取 session-id
#         res_session_id = util.net.get(
#             "https://www.amazon.co.jp/s?k=cat",
#             headers={
#                 "Content-Type": "text/html;charset=UTF-8",
#                 "User-Agent": config.user_agent,
#                 "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
#                 "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
#                 "Accept-Encoding": "gzip, deflate, br, zstd",
#             },
#             stream=True
#         )
#
#         # 获取 anti-csrftoken-a2z
#         csrf_token = ''
#         try:
#             match = re.search(r'&quot;anti-csrftoken-a2z&quot;:&quot;(.*?)&quot;', res_session_id.text)
#             if match:
#                 source_csrf_token = match.group(1)
#
#                 res_rendered_address_selections = util.net.get(
#                     "https://www.amazon.co.jp/portal-migration/hz/glow/get-rendered-address-selections?deviceType=desktop&pageType=Search&storeContext=NoStoreName&actionSource=desktop-modal",
#                     headers={
#                         "Content-Type": "text/html;charset=UTF-8",
#                         "Referer": "https://www.amazon.co.jp/s?k=cat",
#                         "User-Agent": config.user_agent,
#                         "Cookie": res_session_id.headers.get("set-cookie"),
#                         "anti-csrftoken-a2z": source_csrf_token,
#                     }
#                 )
#
#                 match = re.search(r'CSRF_TOKEN\s*:\s*"([^"]+)"', res_rendered_address_selections.text)
#                 if match:
#                     csrf_token = match.group(1)
#         except:
#             pass
#
#         # 再获取 ubid-acbjp
#         res_ubid = util.net.post(
#             "https://www.amazon.co.jp/portal-migration/hz/glow/address-change?actionSource=glow",
#             json_data={
#                 "locationType": "LOCATION_INPUT",
#                 "zipCode": "169-0074",
#                 "deviceType": "web",
#                 "storeContext": "hpc",
#                 "pageType": "Detail",
#                 "actionSource": "glow"
#             },
#             headers={
#                 "Content-Type": "application/json",
#                 "User-Agent": config.user_agent,
#                 "Cookie": res_session_id.headers.get("set-cookie"),
#                 # 这个参数是用来获取该接口返回的 json 数据的，如果不传此参数，可以正常获取包含地址的 Cookie 数据，但服务器会返回空
#                 "anti-csrftoken-a2z": csrf_token,
#             },
#             stream=True
#         )
#
#         session_id = res_session_id.cookies.get("session-id", "")
#         ubid = res_ubid.cookies.get("ubid-acbjp", "")
#
#         if len(session_id) == 0 or len(ubid) == 0:
#             raise Exception("Response is ok, but values are empty")
#         else:
#             try:
#                 addr_data = res_ubid.json().get("address", {})
#                 country = util.region.country_code_to_chinese(addr_data.get("countryCode", ""))
#                 state = addr_data.get("state", "")
#                 city = addr_data.get("city", "")
#                 district = addr_data.get("district", "")
#                 print(f"当前选择的配送地址：{country} {state} {city} {district}", "green")
#             except:
#                 pass
#
#             return "ubid-acbjp=" + ubid + "; session-id=" + session_id
#     except Exception as e:
#         print(f"Cookie 获取失败: {e}", "red")
#         print("将使用备用的 Cookie 数据。", "orange")
#
#         # 从浏览器开发者工具中拿到的 Cookie 数据，有效期一年
#         return "ubid-acbjp=355-5685452-2837352; session-id=357-7564356-4927846"


# if __name__ == '__main__':
#     # print(a())
#
#     res_session_id = util.net.get(
#         "https://www.amazon.co.jp/s?k=cat",
#         # "https://www.google.com",
#         headers={
#             "Content-Type": "text/html;charset=UTF-8",
#             "User-Agent": config.user_agent,
#             "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
#             "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
#             "Accept-Encoding": "gzip, deflate, br, zstd",
#         },
#         stream=True
#     )
#     print(res_session_id.text)


def test():
    with open("C:\\Users\\ceneax\\Desktop\\1.txt", 'r', encoding='utf-8') as f:
        b = f.read()
    soup = BeautifulSoup(b, "html.parser")

    # elements = soup.select('div[cel_widget_id*="MAIN-SEARCH_RESULTS-"]')
    elements = soup.select('div[data-component-type="s-search-result"]')
    valuable_els = []
    today = date.today()

    print(len(elements))

    for el in elements:
        delivery_message = el.select_one("div.udm-primary-delivery-message span.a-text-bold")
        if not delivery_message:
            continue

        print(delivery_message.get_text())

        match = re.search(r'(\d{1,2})月(\d{1,2})日', delivery_message.get_text())
        if not match:
            continue

        month, day = int(match.group(1)), int(match.group(2))

        try:
            # 1. 先假设是今年
            target_date = date(today.year, month, day)

            # 2. 【新增】如果今年的这个日期已经过了，那就加一年（算作明年的日期）
            if target_date < today:
                target_date = target_date.replace(year=today.year + 1)

            offset = (target_date - today).days
            print(offset)
            if 8 <= offset <= 30:
                valuable_els.append(el)
        except ValueError:
            continue  # 忽略无效日期，如2月30日

    if len(valuable_els) < 8:
        print(f'关键词： 匹配到 {len(valuable_els)} 个，不符合要求，忽略。', 'orange')
        return None

    kw_img = None
    products = []

    for ve in valuable_els:
        img = None
        title = ''
        price_symbol = None
        price = None
        buy_number = None

        product_asin = ve['data-asin'] if 'data-asin' in ve.attrs else None
        if product_asin and len(product_asin.strip()) > 0:
            img_el = ve.find('img', class_='s-image')
            img_result = img_el['src'] if img_el and 'src' in img_el.attrs else None
            if img_result is not None and img_result.strip().startswith('http'):
                img = img_result

            title_el = ve.select_one('div[data-cy="title-recipe"]')
            if title_el:
                h2_el = title_el.find('h2')
                if h2_el:
                    title = h2_el.get_text().strip()

            price_el = ve.select_one('a[aria-describedby="price-link"]')
            if price_el:
                p_symbol_el = price_el.find('span', class_='a-price-symbol')
                p_symbol = p_symbol_el.get_text().strip() if p_symbol_el else ''
                if len(p_symbol) > 0:
                    price_symbol = p_symbol

                p_whole_el = price_el.find('span', class_='a-price-whole')
                if p_whole_el:
                    p_whole = p_whole_el.get_text().strip()
                    p_decimal_el = p_whole_el.find('span', class_='a-price-decimal')
                    p_decimal = p_decimal_el.text.strip() if p_decimal_el else ''
                    try:
                        price = int(p_whole.replace(',', '').replace(p_decimal, ''))
                    except Exception:
                        pass

            buy_number_el = ve.select_one('div[data-cy="reviews-block"]')
            if buy_number_el:
                b_number_el = buy_number_el.select_one('span.a-size-base.a-color-secondary')
                b_number_full = b_number_el.get_text().strip() if b_number_el else ''

                # 匹配模式：
                # (\d+)            —— 数字部分（整数）
                # \s*              —— 允许数字和单位之间有空格
                # (万|百万|千万|亿)? —— 单位，可选
                match = re.search(r'过去一个月有(\d+)\s*(万|百万|千万|亿)?\+?', b_number_full)
                if match:
                    try:
                        num_str, unit = match.groups()
                        num = int(num_str)

                        # 单位换算表
                        unit_map = {
                            None: 1,
                            '万': 10_000,
                            '百万': 1_000_000,
                            '千万': 10_000_000,
                            '亿': 100_000_000
                        }

                        factor = unit_map.get(unit, 1)
                        buy_number = int(num * factor)
                    except Exception:
                        pass

            products.append((product_asin.strip(), img, title, price_symbol, price, buy_number))
        else:
            print(f"搜索词 未找到对应商品的 asin 数据", "orange")

        # 如果搜索词需要的图片获取到了，就不用执行下面的代码了
        if kw_img is not None:
            continue

        img_el = ve.find('img', class_='s-image')
        img_result = img_el['src'] if img_el and 'src' in img_el.attrs else None
        if img_result is not None and img_result.strip().startswith('http'):
            kw_img = img_result

        print(f'符合的关键词 {len(valuable_els)}个')


if __name__ == '__main__':
    test()
