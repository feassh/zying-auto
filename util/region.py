# 定义国家代码（ISO 2位字母）到中文名称的映射
COUNTRY_CODE_TO_CHINESE = {
    "JP": "日本",
    "US": "美国",
    "CN": "中国",
    "GB": "英国",
    "FR": "法国",
    "DE": "德国",
    "KR": "韩国",
    "SG": "新加坡",
    "CA": "加拿大",
    "AU": "澳大利亚",
    "IN": "印度",
    "BR": "巴西",
    "RU": "俄罗斯",
    "IT": "意大利",
    "ES": "西班牙",
    "NL": "荷兰",
    "MX": "墨西哥",
    "TH": "泰国",
    "ID": "印度尼西亚",
    "MY": "马来西亚",
    "BE": "比利时",
    "CH": "瑞士",
    "SE": "瑞典",
    "NO": "挪威",
    "DK": "丹麦",
    "FI": "芬兰",
    "AT": "奥地利",
    "PL": "波兰",
    "PT": "葡萄牙",
    "IE": "爱尔兰",
    "PH": "菲律宾",
    "VN": "越南",
    "TR": "土耳其",
    "AE": "阿联酋",
    "SA": "沙特阿拉伯",
    "EG": "埃及",
    "ZA": "南非",
    "NG": "尼日利亚",
    # 可继续添加更多国家...
}


def country_code_to_chinese(country_code):
    # 统一转为大写，避免大小写问题
    code_upper = country_code.upper()
    return COUNTRY_CODE_TO_CHINESE.get(code_upper, "未知国家")
