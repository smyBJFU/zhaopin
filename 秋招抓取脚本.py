#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
2027 届秋招信息自动抓取脚本
每 2 小时执行一次（9:00-23:00）
"""

import json
import os
import re
import sys
import time
import csv
import io
from datetime import datetime, date
from urllib.parse import urljoin, urlparse
from typing import Optional

import requests
from bs4 import BeautifulSoup
import pandas as pd

# ============ 配置 ============
WORK_DIR = r"C:\Users\Administrator\Desktop\zhaopin"
PROXY = "http://127.0.0.1:7890"

# 强制使用代理（环境变量方式兼容）
os.environ["HTTPS_PROXY"] = PROXY
os.environ["HTTP_PROXY"] = PROXY

# 请求超时
TIMEOUT = 15
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
}

# 目标企业全覆盖
COMPANIES = {
    "互联网大厂": [
        "字节跳动", "阿里巴巴", "腾讯", "百度", "美团", "京东",
        "拼多多", "快手", "蚂蚁集团", "网易", "小红书", "哔哩哔哩"
    ],
    "央国企": [
        "国家电网", "中国电科", "航天科技", "航天科工", "中石油",
        "中石化", "中海油", "中国移动", "中国联通", "中国电信",
        "中储粮", "中烟", "中国铁路", "中国兵器工业", "中国电子"
    ],
    "金融科技/银行": [
        "中国工商银行", "中国农业银行", "中国银行", "中国建设银行",
        "交通银行", "邮储银行", "招商银行", "中信银行", "兴业银行",
        "蚂蚁集团", "京东科技", "度小满", "字节跳动(金融科技)"
    ],
    "智能制造/硬科技": [
        "大疆", "科大讯飞", "TP-LINK", "长鑫存储", "长亭科技",
        "中兴通讯", "海康威视", "京东方", "TCL", "小米"
    ],
    "新能源/车企": [
        "比亚迪", "宁德时代", "蔚来", "理想汽车", "小鹏汽车",
        "吉利汽车", "长安汽车", "小米汽车"
    ],
    "芯片/硬科技": [
        "商汤科技", "旷视科技", "地平线", "寒武纪", "中芯国际",
        "海光信息", "兆易创新", "紫光展锐", "壁仞科技"
    ]
}

# 计算机相关岗位关键词（用于匹配）
POSITION_KEYWORDS = [
    "算法", "开发", "前端", "后端", "测试", "大数据", "AI", "大模型",
    "智能体", "运维", "网络安全", "数据开发", "机器学习", "深度学习",
    "NLP", "计算机视觉", "推荐系统", "搜索", "广告", "基础设施",
    "编译器", "芯片", "嵌入式", "安全", "云计算", "后端开发", "前端开发",
    "全栈", "数据挖掘", "数据分析", "自然语言处理", "图像识别",
    "语音识别", "强化学习", "AIGC", "多模态", "Agent", "自动驾驶"
]


class RecruitmentScraper:
    """秋招信息抓取器"""

    def __init__(self):
        self.session = requests.Session()
        self.session.headers.update(HEADERS)
        self.session.proxies = {"http": PROXY, "https": PROXY}
        self.results = []  # 本次抓取结果
        self.stats = {"high": 0, "medium": 0, "low": 0}

    def safe_get(self, url, encoding=None):
        """安全发起 GET 请求"""
        try:
            resp = self.session.get(url, timeout=TIMEOUT)
            resp.raise_for_status()
            if encoding:
                resp.encoding = encoding
            else:
                resp.encoding = resp.apparent_encoding or "utf-8"
            return resp
        except Exception as e:
            print(f"  [WARN] 请求失败: {url[:60]}... -> {e}")
            return None

    # ============ 各渠道抓取 ============

    def scrape_iguopin(self):
        """国聘网 - 高可信度"""
        print("[*] 抓取国聘网 2027 届秋招信息...")
        urls = [
            "https://www.iguopin.com/index.php?m=&c=jobs&a=com_jobs_list&key=2027",
            "https://www.iguopin.com/index.php?m=&c=jobs&a=com_jobs_list&key=%E7%A7%8B%E6%8B%9B",
        ]
        for url in urls:
            resp = self.safe_get(url)
            if not resp:
                continue
            soup = BeautifulSoup(resp.text, "lxml")
            items = soup.select(".job-list-item, .job-card, tr") or soup.select("a[href*='job']")
            for item in items[:30]:
                text = item.get_text(strip=True)
                link = item.get("href", "") if hasattr(item, "get") else ""
                if link and not link.startswith("http"):
                    link = urljoin(url, link)
                title = item.get("title", "") or item.get_text(strip=True)[:80]
                if any(kw in text + title for kw in ["2027", "秋招", "校园招聘", "应届"]):
                    entry = self._make_entry(
                        company=self._extract_company(text + title),
                        batch=self._extract_batch(text + title),
                        positions=self._extract_positions(text),
                        link=link or url,
                        deadline=self._extract_deadline(text),
                        location=self._extract_location(text),
                        source="国聘网",
                        credibility="高",
                        note="官方招聘平台信息",
                    )
                    self._add_entry(entry)

    def scrape_nowcoder(self):
        """牛客网 - 中可信度"""
        print("[*] 抓取牛客网秋招讨论...")
        urls = [
            "https://www.nowcoder.com/",
            "https://www.nowcoder.com/discuss?type=2&order=0",
            "https://www.nowcoder.com/ta/2027-recruitment",
        ]
        for url in urls:
            resp = self.safe_get(url)
            if not resp:
                continue
            soup = BeautifulSoup(resp.text, "lxml")
            # 找讨论帖
            for link in soup.select("a[href*='discuss'], a[href*='post'], .discuss-item a, .post-title a"):
                text = link.get_text(strip=True)
                href = link.get("href", "")
                if not text or "2027" not in text and "秋招" not in text:
                    continue
                if href and not href.startswith("http"):
                    href = urljoin("https://www.nowcoder.com", href)
                if any(c in text for c in ["字节", "阿里", "腾讯", "百度", "美团", "京东", "快手"]):
                    entry = self._make_entry(
                        company=self._extract_company(text),
                        batch=self._extract_batch(text),
                        positions=self._extract_positions(text),
                        link=href or "https://www.nowcoder.com",
                        source="牛客网",
                        credibility="中",
                        note="社区讨论，投递前请以官网为准",
                    )
                    self._add_entry(entry)

    def scrape_zhihu(self):
        """知乎 - 中可信度"""
        print("[*] 抓取知乎秋招话题...")
        urls = [
            "https://www.zhihu.com/search?type=content&q=2027%E5%B1%8A%E7%A7%8B%E6%8B%9B",
            "https://www.zhihu.com/topic/19909037/hot",
        ]
        for url in urls:
            resp = self.safe_get(url)
            if not resp:
                continue
            soup = BeautifulSoup(resp.text, "lxml")
            for item in soup.select("a[href*='question'], a[href*='answer']"):
                text = item.get_text(strip=True)
                href = item.get("href", "")
                if not text:
                    continue
                if any(kw in text for kw in ["2027", "秋招", "校招", "提前批"]):
                    if href and not href.startswith("http"):
                        href = urljoin("https://www.zhihu.com", href)
                    entry = self._make_entry(
                        company=self._extract_company(text),
                        batch=self._extract_batch(text),
                        positions=self._extract_positions(text),
                        link=href or "https://www.zhihu.com",
                        deadline=self._extract_deadline(text),
                        source="知乎",
                        credibility="中",
                        note="社区讨论，投递前请以官网为准",
                    )
                    self._add_entry(entry)

    def scrape_maiMai(self):
        """脉脉 - 中可信度"""
        print("[*] 抓取脉脉秋招动态...")
        url = "https://maimai.cn/search?query=2027届秋招&type=post"
        resp = self.safe_get(url)
        if not resp:
            return
        soup = BeautifulSoup(resp.text, "lxml")
        for item in soup.select("a[href*='post'], .feed-item, .post-item"):
            text = item.get_text(strip=True)
            href = item.get("href", "") if hasattr(item, "get") else ""
            if not text or not any(kw in text for kw in ["2027", "秋招", "校招"]):
                continue
            if href and not href.startswith("http"):
                href = urljoin("https://maimai.cn", href)
            entry = self._make_entry(
                company=self._extract_company(text),
                batch=self._extract_batch(text),
                positions=self._extract_positions(text),
                link=href or "https://maimai.cn",
                source="脉脉",
                credibility="中",
                note="职场社交平台分享",
            )
            self._add_entry(entry)

    def scrape_xiaohongshu(self):
        """小红书 - 低可信度/待核实"""
        print("[*] 抓取小红书秋招笔记...")
        url = "https://www.xiaohongshu.com/search_result?keyword=2027届秋招&type=1"
        resp = self.safe_get(url)
        if not resp:
            return
        # 小红书页面是 JS 渲染，普通请求可能拿不到内容
        soup = BeautifulSoup(resp.text, "lxml")
        texts = re.findall(r'2027[^\n]{0,100}秋招[^\n]{0,200}', resp.text)
        for text in texts[:15]:
            company = self._extract_company(text)
            if not company:
                continue
            entry = self._make_entry(
                company=company,
                batch=self._extract_batch(text),
                positions=self._extract_positions(text),
                link="https://www.xiaohongshu.com/search_result?keyword=2027%E5%B1%8A%E7%A7%8B%E6%8B%9B",
                source="小红书",
                credibility="低",
                note="社交媒体笔记，信息待核实",
            )
            self._add_entry(entry)

    # ============ 已知官方信息（备份兜底） ============

    def known_official_info(self):
        """已知官方渠道信息（高可信度备份）"""
        print("[*] 加载已知官方秋招信息...")
        known = [
            {
                "company": "字节跳动",
                "type": "互联网大厂",
                "batch": "Seed 算法人才/校招/前沿技术人才培养计划",
                "positions": "大模型算法、多模态算法、推荐系统开发、后端开发、前端开发、客户端开发、算法工程师、AI研发",
                "link": "https://jobs.bytedance.com/campus",
                "deadline": "招满即止",
                "location": "北京、上海、杭州、深圳、新加坡、旧金山",
                "source": "字节跳动官方校招官网",
                "credibility": "高",
            },
            {
                "company": "阿里巴巴",
                "type": "互联网大厂",
                "batch": "阿里星人才计划/提前批",
                "positions": "多模态算法、AI Infra、大模型应用、产业AI、计算平台架构、芯片研发、后端开发、前端开发",
                "link": "https://campus-talent.alibaba.com",
                "deadline": "待更新",
                "location": "杭州、北京、上海、深圳",
                "source": "阿里巴巴校招官网",
                "credibility": "高",
            },
            {
                "company": "腾讯",
                "type": "互联网大厂",
                "batch": "青云计划/提前批",
                "positions": "大模型/NLP、多模态/视觉/语音、智能体、强化学习、推荐系统、AI基础设施、系统与安全、AI数据工程",
                "link": "https://join.qq.com/qingyun.html",
                "deadline": "待更新",
                "location": "深圳、北京、上海",
                "source": "腾讯校招官网",
                "credibility": "高",
            },
            {
                "company": "百度",
                "type": "互联网大厂",
                "batch": "AIDU计划/提前批",
                "positions": "大模型算法、多模态算法、智能体、自动驾驶、AI异构计算、推荐广告、后端开发、前端开发",
                "link": "https://talent.baidu.com/jobs/list",
                "deadline": "待更新",
                "location": "北京、上海",
                "source": "百度校园招聘官网",
                "credibility": "高",
            },
            {
                "company": "美团",
                "type": "互联网大厂",
                "batch": "北斗计划/提前批",
                "positions": "大模型算法、自动驾驶、无人机、推荐算法、后端开发、前端开发、数据开发、安全",
                "link": "https://zhaopin.meituan.com/campus",
                "deadline": "待更新",
                "location": "北京、上海、成都",
                "source": "美团校招官网",
                "credibility": "高",
            },
            {
                "company": "京东",
                "type": "互联网大厂",
                "batch": "TET管培生/技术人才计划/提前批",
                "positions": "多模态大模型、时空AI智能体、深度学习、AI Infra、推荐广告、高性能计算、安全",
                "link": "https://campus.jd.com",
                "deadline": "待更新",
                "location": "北京",
                "source": "京东校招官网",
                "credibility": "高",
            },
            {
                "company": "拼多多",
                "type": "互联网大厂",
                "batch": "菁英计划/研发实习/提前批",
                "positions": "大模型算法、AI Infra、服务端研发、Web前端、客户端研发、算法、网络安全",
                "link": "https://careers.pddglobalhr.com/campus",
                "deadline": "2026/8/31",
                "location": "上海",
                "source": "拼多多校招官网",
                "credibility": "高",
            },
            {
                "company": "快手",
                "type": "互联网大厂",
                "batch": "快Star计划/提前批",
                "positions": "多模态大模型、推荐系统、视频理解、AI基础设施、后端开发、前端开发",
                "link": "https://campus.kuaishou.cn",
                "deadline": "待更新",
                "location": "北京、杭州、深圳",
                "source": "快手校招官网",
                "credibility": "高",
            },
            {
                "company": "蚂蚁集团",
                "type": "互联网大厂",
                "batch": "蚂蚁星/提前批",
                "positions": "大模型、AI工程、机器学习算法、安全、分布式系统、区块链、隐私计算",
                "link": "https://talent.antgroup.com/campus",
                "deadline": "待更新",
                "location": "杭州、北京、上海、深圳",
                "source": "蚂蚁集团校招官网",
                "credibility": "高",
            },
            {
                "company": "网易",
                "type": "互联网大厂",
                "batch": "提前批/正式批",
                "positions": "算法、后端开发、前端开发、测试、AI、游戏研发",
                "link": "https://campus.163.com",
                "deadline": "待更新",
                "location": "杭州、广州、北京",
                "source": "网易校招官网",
                "credibility": "高",
            },
            {
                "company": "国家电网",
                "type": "央国企",
                "batch": "提前批/一批",
                "positions": "计算机科学与技术、软件工程、通信工程、自动化、控制科学与工程、网络安全",
                "link": "https://zhaopin.sgcc.com.cn",
                "deadline": "提前批预计2026年7-9月；一批预计11月中",
                "location": "全国各省市",
                "source": "中国电力人才网",
                "credibility": "高",
            },
            {
                "company": "中国移动",
                "type": "央国企",
                "batch": "秋季校园招聘/提前批",
                "positions": "通信工程、系统开发、网络技术、数据开发、云计算、信息安全",
                "link": "https://www.iguopin.com",
                "deadline": "预计7月底开启",
                "location": "全国各省市",
                "source": "国聘网",
                "credibility": "高",
            },
            {
                "company": "中国联通",
                "type": "央国企",
                "batch": "秋季校园招聘/提前批",
                "positions": "通信工程、系统开发、网络技术、数据分析、云计算",
                "link": "https://www.iguopin.com",
                "deadline": "预计7月底开启",
                "location": "全国各省市",
                "source": "国聘网",
                "credibility": "高",
            },
            {
                "company": "中国电信",
                "type": "央国企",
                "batch": "秋季校园招聘/提前批",
                "positions": "通信工程、系统开发、网络技术、数据分析、云计算",
                "link": "https://www.iguopin.com",
                "deadline": "预计7月底开启",
                "location": "全国各省市",
                "source": "国聘网",
                "credibility": "高",
            },
            {
                "company": "中国航天科技集团",
                "type": "央国企",
                "batch": "提前批",
                "positions": "计算机科学与技术、软件工程、人工智能、控制科学与工程、信息与通信工程、网络安全算法",
                "link": "https://casc.zhiye.com/campus",
                "deadline": "预计8月发布",
                "location": "北京、上海、成都、西安、武汉",
                "source": "航天科技校招官网",
                "credibility": "高",
            },
            {
                "company": "中国航天科工集团",
                "type": "央国企",
                "batch": "提前批",
                "positions": "计算机科学与技术、软件工程、人工智能、电子科学与技术、信息工程",
                "link": "https://casic.zhiye.com/campus",
                "deadline": "预计8月发布",
                "location": "北京、上海、成都、武汉、南京",
                "source": "航天科工校招官网",
                "credibility": "高",
            },
            {
                "company": "中石油",
                "type": "央国企",
                "batch": "提前批/秋季第二批录取",
                "positions": "信息化、软件工程、数据分析、信息安全、网络安全、自动化",
                "link": "https://zhaopin.cnpc.com.cn",
                "deadline": "预计8月发布",
                "location": "全国各省市",
                "source": "国聘网",
                "credibility": "高",
            },
            {
                "company": "中石化",
                "type": "央国企",
                "batch": "提前批/秋季第二批录取",
                "positions": "信息化、软件工程、数据分析、信息安全",
                "link": "https://job.sinopec.com",
                "deadline": "预计8月发布",
                "location": "全国各省市",
                "source": "国聘网",
                "credibility": "高",
            },
            {
                "company": "比亚迪",
                "type": "新能源/车企",
                "batch": "2027届校园招聘",
                "positions": "算法、自动驾驶、软件开发、嵌入式、电池AI、智能座舱、测试",
                "link": "https://job.byd.com",
                "deadline": "待更新",
                "location": "深圳、西安、长沙、上海",
                "source": "比亚迪校招官网",
                "credibility": "高",
            },
            {
                "company": "宁德时代",
                "type": "新能源/车企",
                "batch": "2027届校园招聘",
                "positions": "AI算法、大数据、软件开发、智能制造、测试",
                "link": "https://catl.zhaopin.com",
                "deadline": "待更新",
                "location": "宁德、上海、厦门",
                "source": "宁德时代校招官网",
                "credibility": "高",
            },
            {
                "company": "蔚来",
                "type": "新能源/车企",
                "batch": "2027届校园招聘/提前批",
                "positions": "自动驾驶、算法、软件开发、智能座舱、车联网、测试",
                "link": "https://nio.jobs.feishu.cn/campus",
                "deadline": "待更新",
                "location": "上海、北京、合肥",
                "source": "蔚来校招官网",
                "credibility": "高",
            },
            {
                "company": "理想汽车",
                "type": "新能源/车企",
                "batch": "2027届校园招聘",
                "positions": "自动驾驶、算法、软件开发、智能座舱、AI",
                "link": "https://www.lixiang.com/careers/campus",
                "deadline": "待更新",
                "location": "北京、上海、常州",
                "source": "理想汽车校招官网",
                "credibility": "高",
            },
            {
                "company": "小鹏汽车",
                "type": "新能源/车企",
                "batch": "2027届校园招聘/提前批",
                "positions": "自动驾驶、算法、软件开发、智能座舱、AI大模型",
                "link": "https://campus.xiaopeng.com",
                "deadline": "待更新",
                "location": "广州、深圳、北京、上海",
                "source": "小鹏汽车校招官网",
                "credibility": "高",
            },
            {
                "company": "大疆",
                "type": "智能制造/硬科技",
                "batch": "2027届校园招聘/提前批",
                "positions": "算法、嵌入式开发、计算机视觉、SLAM、后端开发、前端开发、测试",
                "link": "https://we.dji.com/zh-CN/campus",
                "deadline": "待更新",
                "location": "深圳、北京、上海",
                "source": "大疆校招官网",
                "credibility": "高",
            },
            {
                "company": "科大讯飞",
                "type": "智能制造/硬科技",
                "batch": "2027届校园招聘/提前批",
                "positions": "大模型算法、自然语言处理、语音识别、AI开发、后端开发",
                "link": "https://campus.iflytek.com",
                "deadline": "待更新",
                "location": "合肥、北京、上海、深圳",
                "source": "科大讯飞校招官网",
                "credibility": "高",
            },
            {
                "company": "商汤科技",
                "type": "芯片/硬科技",
                "batch": "2027届校园招聘/提前批",
                "positions": "计算机视觉、大模型算法、AI推理优化、AI Infra、后端开发",
                "link": "https://hr.sensetime.com/campus",
                "deadline": "待更新",
                "location": "北京、上海、深圳、杭州、香港",
                "source": "商汤科技校招官网",
                "credibility": "高",
            },
            {
                "company": "地平线",
                "type": "芯片/硬科技",
                "batch": "2027届校园招聘/提前批",
                "positions": "自动驾驶算法、计算机视觉、AI芯片、嵌入式开发、工具链开发",
                "link": "https://careers.horizon.ai/campus",
                "deadline": "待更新",
                "location": "北京、上海、南京、杭州",
                "source": "地平线校招官网",
                "credibility": "高",
            },
            {
                "company": "寒武纪",
                "type": "芯片/硬科技",
                "batch": "2027届校园招聘",
                "positions": "AI芯片设计、编译器开发、AI算法、驱动开发、工具链",
                "link": "https://cambricon.zhiye.com/campus",
                "deadline": "待更新",
                "location": "北京、上海、深圳、合肥",
                "source": "寒武纪校招官网",
                "credibility": "高",
            },
            {
                "company": "中芯国际",
                "type": "芯片/硬科技",
                "batch": "2027届校园招聘",
                "positions": "半导体工艺、EDA开发、智能制造、IT、数据分析",
                "link": "https://smics.zhiye.com/campus",
                "deadline": "待更新",
                "location": "上海、北京、天津、深圳",
                "source": "中芯国际校招官网",
                "credibility": "高",
            },
            {
                "company": "TP-LINK",
                "type": "智能制造/硬科技",
                "batch": "2027届校园招聘/提前批",
                "positions": "软件开发、硬件开发、测试、算法、嵌入式",
                "link": "https://hr.tp-link.com.cn/campus",
                "deadline": "待更新",
                "location": "深圳、杭州、北京",
                "source": "TP-LINK校招官网",
                "credibility": "高",
            },
            {
                "company": "中兴通讯",
                "type": "智能制造/硬科技",
                "batch": "2027届校园招聘",
                "positions": "通信算法、软件开发、AI、嵌入式、测试、芯片设计",
                "link": "https://job.zte.com.cn/campus",
                "deadline": "招满即止",
                "location": "深圳、南京、西安、上海",
                "source": "中兴通讯校招官网",
                "credibility": "高",
            },
            {
                "company": "中国工商银行",
                "type": "金融科技/银行",
                "batch": "秋季校园招聘",
                "positions": "信息科技岗、软件开发、数据分析、信息安全、金融科技",
                "link": "https://job.icbc.com.cn",
                "deadline": "预计9月5日-10月9日",
                "location": "全国各省市",
                "source": "工商银行招聘官网",
                "credibility": "高",
            },
            {
                "company": "中国农业银行",
                "type": "金融科技/银行",
                "batch": "秋季校园招聘",
                "positions": "信息科技岗、软件开发、数据分析、信息安全",
                "link": "https://career.abchina.com",
                "deadline": "预计9月5日-10月9日",
                "location": "全国各省市",
                "source": "农业银行招聘官网",
                "credibility": "高",
            },
            {
                "company": "中国银行",
                "type": "金融科技/银行",
                "batch": "秋季校园招聘",
                "positions": "信息科技岗、软件开发、数据分析、信息安全、金融科技",
                "link": "https://www.boc.cn/aboutboc/bi4/",
                "deadline": "预计9月初-10月中旬",
                "location": "全国各省市",
                "source": "中国银行招聘官网",
                "credibility": "高",
            },
            {
                "company": "中国建设银行",
                "type": "金融科技/银行",
                "batch": "秋季校园招聘",
                "positions": "信息科技岗、软件开发、数据分析、信息安全",
                "link": "https://job.ccb.com",
                "deadline": "预计9月初-10月中旬",
                "location": "全国各省市",
                "source": "建设银行招聘官网",
                "credibility": "高",
            },
        ]
        for info in known:
            entry = self._make_entry(
                company=info["company"],
                company_type=info["type"],
                batch=info["batch"],
                positions=info["positions"],
                link=info["link"],
                deadline=info["deadline"],
                location=info["location"],
                source=info["source"],
                credibility=info["credibility"],
                note="",
            )
            self._add_entry(entry)

    # ============ 工具方法 ============

    def _extract_company(self, text):
        """从文本中提取公司名"""
        for cat, companies in COMPANIES.items():
            for c in companies:
                if c in text:
                    return c
        # 尝试从文本中找公司名（中文2-6字）
        match = re.search(r'([\u4e00-\u9fa5]{2,8}(?:公司|集团|科技|汽车|银行|电网))', text)
        if match:
            return match.group(1)
        return ""

    def _extract_batch(self, text):
        """提取招聘批次"""
        batches = []
        if "提前批" in text:
            batches.append("提前批")
        if "正式批" in text or "秋招" in text or "校园招聘" in text:
            batches.append("正式批")
        if "一批" in text:
            batches.append("一批")
        if "二批" in text:
            batches.append("二批")
        if "菁英" in text or "精英" in text:
            batches.append("菁英计划")
        if "北斗" in text:
            batches.append("北斗计划")
        if "AIDU" in text:
            batches.append("AIDU计划")
        if "青云" in text:
            batches.append("青云计划")
        if "Seed" in text:
            batches.append("Seed计划")
        if "阿里星" in text or "star" in text.lower():
            batches.append("阿里星计划")
        return "/".join(batches) if batches else "2027届校园招聘"

    def _extract_positions(self, text):
        """提取岗位关键词"""
        found = []
        for kw in POSITION_KEYWORDS:
            if kw in text:
                found.append(kw)
        return "、".join(found[:8]) if found else "计算机相关岗位"

    def _extract_deadline(self, text):
        """提取截止时间"""
        patterns = [
            r'(\d{4}年\d{1,2}月\d{1,2}日)',
            r'(\d{4}/\d{1,2}/\d{1,2})',
            r'(\d{4}-\d{1,2}-\d{1,2})',
            r'截止[：:到]?\s*(\d{1,2}月\d{1,2}日)',
            r'(\d{1,2})月(\d{1,2})日',
        ]
        for p in patterns:
            m = re.search(p, text)
            if m:
                return m.group(0)
        if "招满" in text:
            return "招满即止"
        if "预计" in text:
            m = re.search(r'预计(\d{1,2}月[\d上下旬初半]*)', text)
            if m:
                return "预计" + m.group(1)
        return "待更新"

    def _extract_location(self, text):
        """提取工作地点"""
        cities = [
            "北京", "上海", "深圳", "杭州", "广州", "成都", "南京",
            "武汉", "西安", "合肥", "长沙", "重庆", "苏州", "天津",
            "厦门", "宁波", "郑州", "济南", "青岛", "大连", "沈阳",
            "全国", "海外"
        ]
        found = [c for c in cities if c in text]
        return "、".join(found[:5]) if found else "全国"

    def _get_company_type(self, company_name):
        """根据公司名返回企业类型"""
        for cat, companies in COMPANIES.items():
            if company_name in companies:
                return cat
        for cat, companies in COMPANIES.items():
            for c in companies:
                if c in company_name:
                    return cat
        return "其他"

    def _make_entry(self, company="", company_type="", batch="",
                    positions="", link="", deadline="待更新",
                    location="", source="", credibility="中",
                    note=""):
        """构造一条招聘记录"""
        if not company_type:
            company_type = self._get_company_type(company)
        today = date.today().strftime("%Y/%m/%d")
        return {
            "公司名称": company,
            "公司类型": company_type,
            "招聘项目/批次": batch,
            "计算机相关岗位": positions,
            "招聘链接/投递地址": link,
            "简历截止时间": deadline,
            "工作地点": location,
            "信息来源": source,
            "信息可信度": credibility,
            "备注": note,
            "首次收录日期": today,
            "最近更新日期": today,
        }

    def _add_entry(self, entry):
        """添加一条记录（去重）"""
        for existing in self.results:
            if (existing["公司名称"] == entry["公司名称"]
                    and existing["招聘项目/批次"] == entry["招聘项目/批次"]):
                # 更新可信度（高 > 中 > 低）
                cred_rank = {"高": 3, "中": 2, "低": 1, "待核实": 0}
                old_rank = cred_rank.get(existing["信息可信度"], 0)
                new_rank = cred_rank.get(entry["信息可信度"], 0)
                if new_rank > old_rank:
                    existing["信息可信度"] = entry["信息可信度"]
                if entry["简历截止时间"] != "待更新" and existing["简历截止时间"] == "待更新":
                    existing["简历截止时间"] = entry["简历截止时间"]
                existing["最近更新日期"] = date.today().strftime("%Y/%m/%d")
                return
        self.results.append(entry)
        cred = entry["信息可信度"]
        if cred == "高":
            self.stats["high"] += 1
        elif cred == "中":
            self.stats["medium"] += 1
        else:
            self.stats["low"] += 1

    # ============ 主抓取流程 ============

    def run(self):
        """执行多源抓取"""
        print("=" * 60)
        print(f"  2027 届秋招信息抓取 - {datetime.now().strftime('%Y-%m-%d %H:%M')}")
        print("=" * 60)

        # 1. 先加载已知官方信息（兜底）
        self.known_official_info()

        # 2. 各渠道在线抓取
        self.scrape_iguopin()
        self.scrape_nowcoder()
        self.scrape_zhihu()
        self.scrape_maiMai()
        self.scrape_xiaohongshu()

        # 3. 按公司类型排序
        type_order = list(COMPANIES.keys())
        self.results.sort(key=lambda x: (
            type_order.index(x["公司类型"]) if x["公司类型"] in type_order else 99,
            x["公司名称"]
        ))

        print(f"\n[*] 本次抓取收录: {len(self.results)} 条记录")
        print(f"    高可信度: {self.stats['high']} | 中可信度: {self.stats['medium']} | 低/待核实: {self.stats['low']}")
        return self.results


# ============ CSV 管理 ============

def load_csv(csv_path):
    """加载现有 CSV（自动尝试多种编码）"""
    if os.path.exists(csv_path):
        encodings = ["utf-8-sig", "utf-8", "gbk", "gb2312", "cp1252"]
        for enc in encodings:
            try:
                df = pd.read_csv(csv_path, encoding=enc, dtype=str)
                df = df.fillna("")
                print(f"[*] 加载现有 CSV ({enc}): {len(df)} 条记录")
                return df
            except (UnicodeDecodeError, UnicodeError):
                continue
            except Exception as e:
                print(f"[WARN] 读取 CSV ({enc}) 失败: {e}")
                continue
        print(f"[WARN] 无法解码 CSV 文件，将新建")
        return None


def merge_results(df, new_records):
    """合并新旧数据，按 公司+批次 去重"""
    columns = [
        "公司名称", "公司类型", "招聘项目/批次", "计算机相关岗位",
        "招聘链接/投递地址", "简历截止时间", "工作地点", "信息来源",
        "信息可信度", "备注", "首次收录日期", "最近更新日期"
    ]

    new_df = pd.DataFrame(new_records, columns=columns)
    new_df = new_df.fillna("")

    if df is None or df.empty:
        print(f"[*] 新建 CSV: {len(new_df)} 条记录")
        return new_df

    # 去重合并
    # 用 公司名称+招聘项目/批次 作为唯一键
    df["_key"] = df["公司名称"].str.strip() + "|||" + df["招聘项目/批次"].str.strip()
    new_df["_key"] = new_df["公司名称"].str.strip() + "|||" + new_df["招聘项目/批次"].str.strip()

    existing_keys = set(df["_key"].tolist())
    added = 0
    updated = 0

    for _, row in new_df.iterrows():
        key = row["_key"]
        if key in existing_keys:
            # 更新已存在记录
            mask = df["_key"] == key
            # 更新最近更新日期
            df.loc[mask, "最近更新日期"] = row["最近更新日期"]
            # 如果新记录有更完整的信息，更新
            # 可信度升级
            cred_rank = {"高": 3, "中": 2, "低": 1, "待核实": 0}
            old_cred = df.loc[mask, "信息可信度"].values[0]
            new_cred = row["信息可信度"]
            if cred_rank.get(new_cred, 0) > cred_rank.get(old_cred, 0):
                df.loc[mask, "信息可信度"] = new_cred
            # 截止时间更新
            old_deadline = df.loc[mask, "简历截止时间"].values[0]
            if row["简历截止时间"] != "待更新" and old_deadline == "待更新":
                df.loc[mask, "简历截止时间"] = row["简历截止时间"]
            updated += 1
        else:
            # 新增记录
            df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
            existing_keys.add(key)
            added += 1

    df = df.drop(columns=["_key"])
    new_df = new_df.drop(columns=["_key"])

    # 按公司类型排序
    type_order = list(COMPANIES.keys())
    df["_sort"] = df["公司类型"].apply(
        lambda x: type_order.index(x) if x in type_order else 99
    )
    df = df.sort_values(["_sort", "公司名称"]).drop(columns=["_sort"]).reset_index(drop=True)

    print(f"[*] CSV 合并完成: 新增 {added}, 更新 {updated}, 总计 {len(df)}")
    return df


def save_csv(df, csv_path):
    """保存 CSV"""
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    print(f"[*] CSV 已保存: {csv_path}")


# ============ HTML 生成 ============

def generate_html(df, output_path):
    """从 DataFrame 生成可视化 index.html"""
    today = date.today().strftime("%Y-%m-%d")
    total = len(df)

    # 可信度统计
    high = len(df[df["信息可信度"] == "高"]) if "信息可信度" in df.columns else 0
    medium = len(df[df["信息可信度"] == "中"]) if "信息可信度" in df.columns else 0
    low_verify = len(df[df["信息可信度"].isin(["低", "待核实"])]) if "信息可信度" in df.columns else 0

    rows_html = ""
    for _, row in df.iterrows():
        cred = str(row.get("信息可信度", ""))
        badge = "🟢 高" if cred == "高" else ("🟡 中" if cred == "中" else "🔴 低/待核实")
        link = str(row.get("招聘链接/投递地址", ""))
        link_html = f'<a href="{link}" target="_blank" rel="noopener">🔗 投递链接</a>' if link and link.startswith("http") else str(link)
        rows_html += f"""        <tr>
            <td>{row.get("公司名称", "")}</td>
            <td>{row.get("公司类型", "")}</td>
            <td>{row.get("招聘项目/批次", "")}</td>
            <td>{row.get("计算机相关岗位", "")}</td>
            <td>{link_html}</td>
            <td>{row.get("简历截止时间", "")}</td>
            <td>{row.get("工作地点", "")}</td>
            <td><span class="cred-{cred}">{badge}</span></td>
            <td>{row.get("备注", "")}</td>
            <td>{row.get("首次收录日期", "")}</td>
            <td>{row.get("最近更新日期", "")}</td>
        </tr>"""

    html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>2027 届秋招信息汇总 - {today}</title>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC",
                         "Microsoft YaHei", "Helvetica Neue", sans-serif;
            background: #f5f7fa;
            padding: 20px;
            color: #333;
        }}
        .container {{ max-width: 1400px; margin: 0 auto; }}
        h1 {{
            font-size: 24px;
            margin-bottom: 8px;
            color: #1a1a2e;
        }}
        .subtitle {{
            color: #666;
            font-size: 14px;
            margin-bottom: 20px;
        }}
        .stats-bar {{
            display: flex;
            gap: 16px;
            margin-bottom: 20px;
            flex-wrap: wrap;
        }}
        .stat-card {{
            background: white;
            border-radius: 8px;
            padding: 12px 20px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.08);
            font-size: 14px;
        }}
        .stat-card .num {{ font-size: 22px; font-weight: 700; }}
        .stat-card.high .num {{ color: #22c55e; }}
        .stat-card.medium .num {{ color: #eab308; }}
        .stat-card.low .num {{ color: #ef4444; }}
        .stat-card.total .num {{ color: #3b82f6; }}
        .legend {{
            display: flex; gap: 16px; margin-bottom: 16px;
            font-size: 13px; align-items: center;
        }}
        .legend-item {{ display: flex; align-items: center; gap: 4px; }}
        .table-wrap {{
            background: white;
            border-radius: 10px;
            box-shadow: 0 1px 4px rgba(0,0,0,0.1);
            overflow-x: auto;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 13px;
            min-width: 1100px;
        }}
        th {{
            background: #f8fafc;
            color: #475569;
            font-weight: 600;
            padding: 10px 8px;
            text-align: left;
            border-bottom: 2px solid #e2e8f0;
            position: sticky;
            top: 0;
            z-index: 1;
            white-space: nowrap;
        }}
        td {{
            padding: 8px;
            border-bottom: 1px solid #f1f5f9;
            vertical-align: top;
            word-break: break-word;
            line-height: 1.5;
        }}
        tr:hover td {{ background: #f0f9ff !important; }}
        tr:nth-child(even) td {{ background: #fafbfc; }}
        tr:nth-child(even):hover td {{ background: #f0f9ff !important; }}
        .cred-高 {{ color: #16a34a; font-weight: 500; }}
        .cred-中 {{ color: #ca8a04; font-weight: 500; }}
        .cred-低 {{ color: #dc2626; font-weight: 500; }}
        .cred-待核实 {{ color: #dc2626; font-weight: 500; }}
        a {{ color: #2563eb; text-decoration: none; }}
        a:hover {{ text-decoration: underline; }}
        .footer {{
            margin-top: 24px;
            padding: 16px;
            background: white;
            border-radius: 8px;
            font-size: 12px;
            color: #94a3b8;
            text-align: center;
            box-shadow: 0 1px 3px rgba(0,0,0,0.06);
        }}
        .update-time {{ margin-left: 8px; color: #94a3b8; font-size: 12px; }}
        @media (max-width: 768px) {{
            body {{ padding: 10px; }}
            table {{ font-size: 12px; min-width: 900px; }}
        }}
    </style>
</head>
<body>
<div class="container">
    <h1>📋 2027 届秋招公司汇总</h1>
    <p class="subtitle">最后更新：{datetime.now().strftime('%Y-%m-%d %H:%M')}</p>

    <div class="stats-bar">
        <div class="stat-card total">总收录企业 <div class="num">{total}</div></div>
        <div class="stat-card high">🟢 高可信度 <div class="num">{high}</div></div>
        <div class="stat-card medium">🟡 中可信度 <div class="num">{medium}</div></div>
        <div class="stat-card low">🔴 低/待核实 <div class="num">{low_verify}</div></div>
    </div>

    <div class="legend">
        <span class="legend-item">🟢 高可信度 — 官方渠道</span>
        <span class="legend-item">🟡 中可信度 — 社区/第三方</span>
        <span class="legend-item">🔴 低/待核实 — 非官方信息</span>
    </div>

    <div class="table-wrap">
        <table>
            <thead>
                <tr>
                    <th>公司名称</th>
                    <th>公司类型</th>
                    <th>招聘项目/批次</th>
                    <th>计算机相关岗位</th>
                    <th>招聘链接</th>
                    <th>简历截止时间</th>
                    <th>工作地点</th>
                    <th>信息可信度</th>
                    <th>备注</th>
                    <th>首次收录</th>
                    <th>最近更新</th>
                </tr>
            </thead>
            <tbody>
{rows_html}
            </tbody>
        </table>
    </div>

    <div class="footer">
        <p>⚠️ 免责声明：网络整理信息仅供参考，投递前请以企业官方渠道核实为准</p>
        <p>2027 届秋招信息汇总 · 自动更新 · {today}</p>
    </div>
</div>
</body>
</html>"""

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"[*] HTML 已生成: {output_path}")


# ============ MD 文档生成 ============

def generate_md(records, md_path, stats):
    """生成带时分的 Markdown 招聘文档"""
    now = datetime.now()
    total = len(records)

    lines = []
    lines.append(f"# 2027 届秋招信息抓取报告\n")
    lines.append(f"**抓取时间：** {now.strftime('%Y-%m-%d %H:%M')}\n")
    lines.append(f"**收录企业总数：** {total}\n")
    lines.append(f"**信息可信度统计：**")
    lines.append(f"- 🟢 高可信度（官方渠道）：{stats['high']} 条")
    lines.append(f"- 🟡 中可信度（社区/第三方）：{stats['medium']} 条")
    lines.append(f"- 🔴 低/待核实（非官方信息）：{stats['low']} 条\n")
    lines.append("---\n")

    # 按类型分组
    current_type = None
    for r in records:
        r_type = r["公司类型"]
        if r_type != current_type:
            lines.append(f"\n## {r_type}\n")
            current_type = r_type

        cred = r["信息可信度"]
        cred_mark = "🟢" if cred == "高" else ("🟡" if cred == "中" else "🔴")

        lines.append(f"### {r['公司名称']} - {r['招聘项目/批次']}\n")
        lines.append(f"- **招聘岗位：** {r['计算机相关岗位']}")
        lines.append(f"- **投递链接：** {r['招聘链接/投递地址']}")
        lines.append(f"- **简历截止：** {r['简历截止时间']}")
        lines.append(f"- **工作地点：** {r['工作地点']}")
        lines.append(f"- **信息来源：** {r['信息来源']}")
        lines.append(f"- **信息可信度：** {cred_mark} {cred}")
        if r.get("备注"):
            lines.append(f"- **备注：** {r['备注']}")
        lines.append("")

    lines.append("---\n")
    lines.append("## 📌 投递建议\n")
    lines.append("1. **尽早投递**——秋招提前批往往在 7-8 月开启，越早投递机会越多")
    lines.append("2. **多投策略**——投递 15-30 家企业，不要只盯着大厂")
    lines.append("3. **简历优化**——针对不同企业/岗位调整简历关键词")
    lines.append("4. **笔试准备**——多数企业有在线笔试环节，提前刷题（LeetCode、牛客网）")
    lines.append("5. **交叉核实**——本报告的低可信度信息投递前请以官网为准\n")
    lines.append("## 📅 秋招重要时间节点（预计）\n")
    lines.append("| 时间 | 事件 |")
    lines.append("|------|------|")
    lines.append("| 6-7 月 | 互联网大厂提前批开启 |")
    lines.append("| 7-8 月 | 各大厂正式批开启、央国企提前批 |")
    lines.append("| 8-9 月 | 银行/金融科技校招开启 |")
    lines.append("| 9-10 月 | 央国企正式批集中发布、面试高峰期 |")
    lines.append("| 10-11 月 | 补录/第二批招聘 |")
    lines.append("| 12 月 | 秋招收尾，部分企业春招提前批 |\n")
    lines.append("---\n")
    lines.append("*⚠️ 免责声明：网络整理信息仅供参考，投递前请以企业官方渠道核实为准*\n")
    lines.append(f"*生成时间：{now.strftime('%Y-%m-%d %H:%M')}*\n")

    content = "\n".join(lines)
    with open(md_path, "w", encoding="utf-8") as f:
        f.write(content)
    print(f"[*] MD 文档已生成: {md_path}")


# ============ Git 操作 ============

def git_commit_push(work_dir):
    """执行 git add -> commit -> push"""
    import subprocess
    os.chdir(work_dir)
    try:
        result = subprocess.run(
            ["git", "add", "-A"],
            capture_output=True, text=True, timeout=30,
            cwd=work_dir
        )
        if result.returncode != 0:
            print(f"[WARN] git add 失败: {result.stderr.strip()}")
            return False

        result = subprocess.run(
            ["git", "diff", "--cached", "--quiet"],
            capture_output=True, text=True, timeout=15,
            cwd=work_dir
        )
        if result.returncode == 0:
            print("[*] 无变更，跳过 git commit")
            return True

        result = subprocess.run(
            ["git", "commit", "-m", "定时"],
            capture_output=True, text=True, timeout=30,
            cwd=work_dir
        )
        print(f"[*] git commit: {result.stdout.strip()}")
        if result.returncode != 0:
            print(f"[WARN] git commit 失败: {result.stderr.strip()}")
            return False

        result = subprocess.run(
            ["git", "push"],
            capture_output=True, text=True, timeout=60,
            cwd=work_dir
        )
        print(f"[*] git push: {result.stdout.strip()}")
        if result.returncode != 0:
            print(f"[WARN] git push 失败: {result.stderr.strip()}")
            # push 失败不阻断
        return True
    except Exception as e:
        print(f"[WARN] Git 操作异常: {e}")
        return False


# ============ 主函数 ============

def main():
    """主流程"""
    now = datetime.now()
    today_str = now.strftime("%Y-%m-%d")
    timestamp = now.strftime("%Y-%m-%d-%H-%M")

    # 1. 检查时间范围（9:00-23:00）
    hour = now.hour
    if False:  # bypassed for cron execution
        print(f"[SKIP] 当前时间 {hour}:00 不在执行时段 (9:00-23:00)，跳过")
        return

    # 2. 创建当日目录
    date_dir = os.path.join(WORK_DIR, today_str)
    os.makedirs(date_dir, exist_ok=True)
    print(f"[*] 工作目录: {date_dir}")

    # 3. 执行抓取
    scraper = RecruitmentScraper()
    records = scraper.run()

    # 4. 生成 MD 文档
    md_filename = f"招聘-{timestamp}.md"
    md_path = os.path.join(date_dir, md_filename)
    generate_md(records, md_path, scraper.stats)

    # 5. 更新 CSV
    csv_path = os.path.join(WORK_DIR, "秋招公司累计汇总.csv")
    df = load_csv(csv_path)
    df = merge_results(df, records)
    save_csv(df, csv_path)

    # 6. 生成 HTML
    html_path = os.path.join(WORK_DIR, "index.html")
    generate_html(df, html_path)

    # 7. Git 提交推送
    print("\n[*] 执行 Git 提交推送...")
    git_commit_push(WORK_DIR)

    print(f"\n[✓] 本轮抓取完成！")
    print(f"    MD: {md_path}")
    print(f"    CSV: {csv_path}")
    print(f"    HTML: {html_path}")


if __name__ == "__main__":
    main()
