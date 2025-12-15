import os
import re

import requests

import config
import util


def a():
    try:
        print("正在尝试获取最新的亚马逊 Cookie...")

        # 先获取 session-id
        res_session_id = util.net.get(
            "https://www.amazon.co.jp/s?k=cat",
            headers={
                "Content-Type": "text/html;charset=UTF-8",
                "User-Agent": config.user_agent,
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
                "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
                "Accept-Encoding": "gzip, deflate, br, zstd",
            },
            stream=True
        )

        # 获取 anti-csrftoken-a2z
        csrf_token = ''
        try:
            match = re.search(r'&quot;anti-csrftoken-a2z&quot;:&quot;(.*?)&quot;', res_session_id.text)
            if match:
                source_csrf_token = match.group(1)

                res_rendered_address_selections = util.net.get(
                    "https://www.amazon.co.jp/portal-migration/hz/glow/get-rendered-address-selections?deviceType=desktop&pageType=Search&storeContext=NoStoreName&actionSource=desktop-modal",
                    headers={
                        "Content-Type": "text/html;charset=UTF-8",
                        "Referer": "https://www.amazon.co.jp/s?k=cat",
                        "User-Agent": config.user_agent,
                        "Cookie": res_session_id.headers.get("set-cookie"),
                        "anti-csrftoken-a2z": source_csrf_token,
                    }
                )

                match = re.search(r'CSRF_TOKEN\s*:\s*"([^"]+)"', res_rendered_address_selections.text)
                if match:
                    csrf_token = match.group(1)
        except:
            pass

        # 再获取 ubid-acbjp
        res_ubid = util.net.post(
            "https://www.amazon.co.jp/portal-migration/hz/glow/address-change?actionSource=glow",
            json_data={
                "locationType": "LOCATION_INPUT",
                "zipCode": "169-0074",
                "deviceType": "web",
                "storeContext": "hpc",
                "pageType": "Detail",
                "actionSource": "glow"
            },
            headers={
                "Content-Type": "application/json",
                "User-Agent": config.user_agent,
                "Cookie": res_session_id.headers.get("set-cookie"),
                # 这个参数是用来获取该接口返回的 json 数据的，如果不传此参数，可以正常获取包含地址的 Cookie 数据，但服务器会返回空
                "anti-csrftoken-a2z": csrf_token,
            },
            stream=True
        )

        session_id = res_session_id.cookies.get("session-id", "")
        ubid = res_ubid.cookies.get("ubid-acbjp", "")

        if len(session_id) == 0 or len(ubid) == 0:
            raise Exception("Response is ok, but values are empty")
        else:
            try:
                addr_data = res_ubid.json().get("address", {})
                country = util.region.country_code_to_chinese(addr_data.get("countryCode", ""))
                state = addr_data.get("state", "")
                city = addr_data.get("city", "")
                district = addr_data.get("district", "")
                print(f"当前选择的配送地址：{country} {state} {city} {district}", "green")
            except:
                pass

            return "ubid-acbjp=" + ubid + "; session-id=" + session_id
    except Exception as e:
        print(f"Cookie 获取失败: {e}", "red")
        print("将使用备用的 Cookie 数据。", "orange")

        # 从浏览器开发者工具中拿到的 Cookie 数据，有效期一年
        return "ubid-acbjp=355-5685452-2837352; session-id=357-7564356-4927846"


if __name__ == '__main__':
    # print(a())

    res_session_id = util.net.get(
        "https://www.amazon.co.jp/s?k=cat",
        # "https://www.google.com",
        headers={
            "Content-Type": "text/html;charset=UTF-8",
            "User-Agent": config.user_agent,
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
            "Accept-Encoding": "gzip, deflate, br, zstd",
        },
        stream=True
    )
    print(res_session_id.text)