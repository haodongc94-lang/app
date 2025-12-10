from dataclasses import dataclass
from typing import Dict, List, Optional
from datetime import datetime
import re
import json
import os
import random
import urllib.request
import time
import sys


@dataclass
class Template:
    id: str
    name: str
    description: str
    fields: List[str]
    body: str
    styles: List[str]


def _today() -> str:
    return datetime.now().strftime("%Y年%m月%d日")


_TRAIN_FILE = "training_data.json"
_LEARNED_FILE = "learned_defaults.json"


def _load_json(path: str):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _save_json(path: str, obj):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)


def list_templates() -> List[Template]:
    t1 = Template(
        id="complaint",
        name="民事起诉状",
        description="用于向人民法院提起民事诉讼",
        fields=["原告姓名", "原告性别", "原告身份证号", "原告地址", "被告姓名", "被告地址", "案由", "诉讼请求", "事实与理由", "法院名称", "日期"],
        body=(
            "{法院名称}\n\n"
            "民事起诉状\n\n"
            "原告：{原告姓名}，{原告性别}，身份证号：{原告身份证号}，住所地：{原告地址}。\n"
            "被告：{被告姓名}，住所地：{被告地址}。\n\n"
            "案由：{案由}。\n\n"
            "诉讼请求：{诉讼请求}。\n\n"
            "事实与理由：{事实与理由}。\n\n"
            "此致\n{法院名称}\n\n"
            "具状人：{原告姓名}\n"
            "日期：{日期}\n"
        ),
        styles=["formal", "neutral", "strict"],
    )
    t2 = Template(
        id="contract",
        name="合同协议书",
        description="双方签订通用合同文本",
        fields=["合同标题", "甲方名称", "乙方名称", "合同标的", "合同期限", "价款与支付", "违约责任", "争议解决", "日期"],
        body=(
            "{合同标题}\n\n"
            "甲方：{甲方名称}\n"
            "乙方：{乙方名称}\n\n"
            "合同标的：{合同标的}\n"
            "合同期限：{合同期限}\n"
            "价款与支付：{价款与支付}\n"
            "违约责任：{违约责任}\n"
            "争议解决：{争议解决}\n\n"
            "签署：\n甲方代表：{甲方名称}\n乙方代表：{乙方名称}\n"
            "日期：{日期}\n"
        ),
        styles=["formal", "neutral"],
    )
    t3 = Template(
        id="power_of_attorney",
        name="授权委托书",
        description="委托他人代为处理相关事务",
        fields=["委托人姓名", "受托人姓名", "委托事项", "委托权限", "委托期限", "日期"],
        body=(
            "授权委托书\n\n"
            "委托人：{委托人姓名}\n"
            "受托人：{受托人姓名}\n\n"
            "委托事项：{委托事项}\n"
            "委托权限：{委托权限}\n"
            "委托期限：{委托期限}\n\n"
            "委托人签名：{委托人姓名}\n"
            "日期：{日期}\n"
        ),
        styles=["formal", "neutral", "strict"],
    )
    t4 = Template(
        id="leave",
        name="请假申请",
        description="员工请假申请文书",
        fields=[
            "申请人姓名",
            "部门",
            "请假类型",
            "请假开始时间",
            "请假结束时间",
            "请假天数",
            "请假事由",
            "审批人",
            "申请日期",
        ],
        body=(
            "请假申请\n\n"
            "申请人：{申请人姓名}\n"
            "部门：{部门}\n\n"
            "请假类型：{请假类型}\n"
            "请假时间：{请假开始时间} 至 {请假结束时间}\n"
            "请假天数：{请假天数} 天\n\n"
            "请假事由：{请假事由}\n\n"
            "审批人：{审批人}\n\n"
            "申请人签名：{申请人姓名}\n"
            "申请日期：{申请日期}\n"
        ),
        styles=["formal", "neutral", "strict"],
    )
    t5 = Template(
        id="meeting_minutes",
        name="会议纪要",
        description="记录会议要点与决议",
        fields=["会议主题", "会议时间", "会议地点", "主持人", "参会人员", "主要议题", "讨论内容", "决议事项", "后续行动", "日期"],
        body=(
            "会议纪要\n\n"
            "会议主题：{会议主题}\n"
            "会议时间：{会议时间}\n"
            "会议地点：{会议地点}\n"
            "主持人：{主持人}\n"
            "参会人员：{参会人员}\n\n"
            "主要议题：{主要议题}\n\n"
            "讨论内容：{讨论内容}\n\n"
            "决议事项：{决议事项}\n\n"
            "后续行动：{后续行动}\n\n"
            "纪要日期：{日期}\n"
        ),
        styles=["formal", "neutral", "strict"],
    )
    t6 = Template(
        id="recommendation_letter",
        name="推荐信",
        description="用于学术或工作推荐",
        fields=["推荐人姓名", "被推荐人姓名", "推荐人单位", "被推荐人背景", "推荐理由", "能力评价", "结语", "日期"],
        body=(
            "推荐信\n\n"
            "推荐人：{推荐人姓名}（{推荐人单位}）\n"
            "被推荐人：{被推荐人姓名}\n\n"
            "背景：{被推荐人背景}\n"
            "推荐理由：{推荐理由}\n"
            "能力评价：{能力评价}\n\n"
            "结语：{结语}\n"
            "日期：{日期}\n"
        ),
        styles=["formal", "neutral", "strict"],
    )
    t7 = Template(
        id="internship_application",
        name="实习申请",
        description="学生/求职者的实习申请文书",
        fields=["申请人姓名", "学校与专业", "实习岗位", "实习单位", "实习时间", "个人优势", "申请理由", "指导老师", "日期"],
        body=(
            "实习申请\n\n"
            "申请人：{申请人姓名}\n"
            "学校与专业：{学校与专业}\n"
            "实习单位与岗位：{实习单位}，{实习岗位}\n"
            "实习时间：{实习时间}\n\n"
            "个人优势：{个人优势}\n"
            "申请理由：{申请理由}\n\n"
            "指导老师：{指导老师}\n"
            "日期：{日期}\n"
        ),
        styles=["formal", "neutral", "strict"],
    )
    t8 = Template(
        id="research_proposal",
        name="研究计划书",
        description="科研课题研究方案",
        fields=["课题名称", "研究背景", "研究目标", "方法与技术路线", "预期成果", "时间安排", "经费预算", "指导老师", "日期"],
        body=(
            "研究计划书\n\n"
            "课题名称：{课题名称}\n\n"
            "研究背景：{研究背景}\n\n"
            "研究目标：{研究目标}\n\n"
            "方法与技术路线：{方法与技术路线}\n\n"
            "预期成果：{预期成果}\n\n"
            "时间安排：{时间安排}\n\n"
            "经费预算：{经费预算}\n\n"
            "指导老师：{指导老师}\n"
            "日期：{日期}\n"
        ),
        styles=["formal", "neutral", "strict"],
    )
    t11 = Template(
        id="project_proposal",
        name="项目立项申请",
        description="项目申报与立项文书",
        fields=["项目名称", "申报单位", "项目背景", "建设目标", "建设内容", "技术方案", "实施计划", "预算与资金来源", "风险与对策", "预期效益", "负责人", "日期"],
        body=(
            "项目立项申请\n\n"
            "项目名称：{项目名称}\n"
            "申报单位：{申报单位}\n\n"
            "项目背景：{项目背景}\n\n"
            "建设目标：{建设目标}\n\n"
            "建设内容：{建设内容}\n\n"
            "技术方案：{技术方案}\n\n"
            "实施计划：{实施计划}\n\n"
            "预算与资金来源：{预算与资金来源}\n\n"
            "风险与对策：{风险与对策}\n\n"
            "预期效益：{预期效益}\n\n"
            "负责人：{负责人}\n"
            "日期：{日期}\n"
        ),
        styles=["formal", "neutral", "strict"],
    )
    t12 = Template(
        id="data_analysis_report",
        name="数据分析报告",
        description="数据分析流程与结论",
        fields=["报告标题", "作者", "数据来源", "清洗与预处理", "统计特征", "建模方法", "评估指标", "结果与可视化", "结论与建议", "日期"],
        body=(
            "数据分析报告\n\n"
            "标题：{报告标题}\n"
            "作者：{作者}\n\n"
            "数据来源：{数据来源}\n\n"
            "清洗与预处理：{清洗与预处理}\n\n"
            "统计特征：{统计特征}\n\n"
            "建模方法：{建模方法}\n\n"
            "评估指标：{评估指标}\n\n"
            "结果与可视化：{结果与可视化}\n\n"
            "结论与建议：{结论与建议}\n\n"
            "日期：{日期}\n"
        ),
        styles=["formal", "neutral", "strict"],
    )
    return [t1, t2, t3, t4, t5, t6, t7, t8, t11, t12]


def _find_template(templates: List[Template], template_id: str) -> Optional[Template]:
    for t in templates:
        if t.id == template_id:
            return t
    return None


def _smart_defaults(template: Template, data: Dict[str, str]) -> Dict[str, str]:
    r = dict(data)
    if "日期" in template.fields and not r.get("日期"):
        r["日期"] = _today()
    learned = _load_json(_LEARNED_FILE)
    ld = learned.get(template.id, {}) if isinstance(learned, dict) else {}
    for f in template.fields:
        if not r.get(f) and f in ld and isinstance(ld[f], str) and ld[f]:
            r[f] = ld[f]
    if template.id == "complaint":
        if not r.get("法院名称"):
            r["法院名称"] = "××人民法院"
        if not r.get("诉讼请求"):
            r["诉讼请求"] = "请求依法判令被告承担相应民事责任"
        if not r.get("事实与理由") and r.get("案由"):
            r["事实与理由"] = f"因{r['案由']}引发纠纷，现依据相关法律提出诉讼"
        if not r.get("原告性别"):
            r["原告性别"] = "男"
    if template.id == "contract":
        if not r.get("合同标题"):
            s = r.get("合同标的") or "合作事宜"
            r["合同标题"] = f"关于{s}之合同协议书"
        if not r.get("争议解决"):
            r["争议解决"] = "双方协商不成的，提交甲方所在地人民法院处理"
        if not r.get("违约责任"):
            r["违约责任"] = "违约方应承担由此产生的全部损失"
    if template.id == "power_of_attorney":
        if not r.get("委托权限"):
            r["委托权限"] = "代为签署相关文件、递交材料、领取文书"
        if not r.get("委托期限"):
            r["委托期限"] = "自本委托书出具之日起至事项办理完毕"
    if template.id == "leave":
        if not r.get("申请日期"):
            r["申请日期"] = _today()
        if not r.get("请假类型"):
            r["请假类型"] = "事假"
        if not r.get("审批人"):
            r["审批人"] = "直属主管"
        if not r.get("请假事由"):
            r["请假事由"] = "因个人事务需处理，特此请假"
        if not r.get("请假天数"):
            r["请假天数"] = "1"
    if template.id == "meeting_minutes":
        if not r.get("日期"):
            r["日期"] = _today()
        if not r.get("后续行动"):
            r["后续行动"] = "责任人明确，按计划推进，定期复盘"
    if template.id == "recommendation_letter":
        if not r.get("结语"):
            r["结语"] = "特此推荐，敬请审阅"
        if not r.get("日期"):
            r["日期"] = _today()
    if template.id == "internship_application":
        if not r.get("申请理由"):
            r["申请理由"] = "希望在实际场景中提升专业能力"
        if not r.get("实习时间"):
            r["实习时间"] = "暑期两个月"
        if not r.get("日期"):
            r["日期"] = _today()
    if template.id == "research_proposal":
        if not r.get("时间安排"):
            r["时间安排"] = "分阶段实施：调研-设计-实验-总结"
        if not r.get("日期"):
            r["日期"] = _today()
    if template.id == "project_proposal":
        if not r.get("预期效益"):
            r["预期效益"] = "提升效率与质量，形成可复制经验"
        if not r.get("日期"):
            r["日期"] = _today()
    if template.id == "data_analysis_report":
        if not r.get("评估指标"):
            r["评估指标"] = "MAE、RMSE、AUC、F1等依任务选择"
        if not r.get("日期"):
            r["日期"] = _today()
    return r


def _apply_style(text: str, style: str) -> str:
    if style not in {"formal", "neutral", "strict"}:
        return text
    rules = []
    if style == "formal":
        rules = [
            (r"(?<!诉讼)请求(?!：)", "恳请"),
            (r"依据", "依照"),
            (r"提交", "谨此提交"),
            (r"违约", "违约行为"),
            (r"处理", "审理处理"),
        ]
        for pat, rep in rules:
            text = re.sub(pat, rep, text)
        reason = None
        m1 = re.search(r"案由：(.+?)。", text)
        if m1:
            reason = m1.group(1)
        m2 = re.search(r"请假事由：(.+?)\n", text)
        if not reason and m2:
            reason = m2.group(1)
        if "民事起诉状" in text:
            pre = f"兹因{reason or '相关纠纷'}，谨此呈请贵院审理。\n\n"
            text = re.sub(r"^(民事起诉状\s*\n+)", r"\1" + pre, text)
        elif "合同协议书" in text:
            pre = "为明确双方权利义务，特订立本协议。\n\n"
            text = re.sub(r"^(合同协议书\s*\n+)", r"\1" + pre, text)
        elif "授权委托书" in text:
            pre = "兹委托受托人依法办理相关事宜。\n\n"
            text = re.sub(r"^(授权委托书\s*\n+)", r"\1" + pre, text)
        elif "请假申请" in text:
            pre = f"兹因{reason or '个人事务'}需处理，谨此申请请假。\n\n"
            text = re.sub(r"^(请假申请\s*\n+)", r"\1" + pre, text)
        return text
    if style == "strict":
        rules = [
            (r"(?<!诉讼)请求(?!：)", "特此请求"),
            (r"依据", "依法律规定"),
            (r"事实与理由：", "事实与法律依据："),
            (r"诉讼请求：", "请求事项："),
            (r"委托事项：", "委托事宜："),
            (r"请假事由：", "事由："),
        ]
        for pat, rep in rules:
            text = re.sub(pat, rep, text)
        if "民事起诉状" in text:
            pre = "经查明，现依法提出如下请求。\n\n"
            text = re.sub(r"^(民事起诉状\s*\n+)", r"\1" + pre, text)
        elif "合同协议书" in text:
            pre = "为规范履约，双方特约如下条款。\n\n"
            text = re.sub(r"^(合同协议书\s*\n+)", r"\1" + pre, text)
        elif "授权委托书" in text:
            pre = "特此授权，受托人按本委托行事。\n\n"
            text = re.sub(r"^(授权委托书\s*\n+)", r"\1" + pre, text)
        elif "请假申请" in text:
            pre = "现依制度申请请假如下。\n\n"
            text = re.sub(r"^(请假申请\s*\n+)", r"\1" + pre, text)
        return text
    return text


def _normalize(text: str) -> str:
    x = re.sub(r"\s+\n", "\n", text)
    x = re.sub(r"\n{3,}", "\n\n", x)
    x = re.sub(r"[。]{2,}", "。", x)
    return x.strip() + "\n"


def generate_document(template_id: str, data: Dict[str, str], style: str = "formal") -> str:
    templates = list_templates()
    t = _find_template(templates, template_id)
    if not t:
        raise ValueError("模板不存在")
    d = _smart_defaults(t, data)
    missing = [f for f in t.fields if f not in d]
    for k in missing:
        d[k] = ""
    text = t.body.format(**d)
    text = _apply_style(text, style)
    return _normalize(text)


def render_preview(template_id: str, data: Dict[str, str], style: str = "formal") -> str:
    return generate_document(template_id, data, style)


def template_fields(template_id: str) -> List[str]:
    t = _find_template(list_templates(), template_id)
    if not t:
        return []
    return t.fields


def template_list() -> List[Dict[str, str]]:
    items = []
    for t in list_templates():
        items.append({"id": t.id, "name": t.name, "description": t.description})
    return items


def record_training(template_id: str, data: Dict[str, str]):
    t = _find_template(list_templates(), template_id)
    if not t:
        return
    rows = []
    for f in t.fields:
        v = (data.get(f) or "").strip()
        if v:
            rows.append({"template_id": template_id, "field": f, "value": v})
    store = _load_json(_TRAIN_FILE)
    if not isinstance(store, dict):
        store = {}
    lst = store.get("rows", [])
    if not isinstance(lst, list):
        lst = []
    lst.extend(rows)
    store["rows"] = lst
    _save_json(_TRAIN_FILE, store)


def run_training() -> Dict[str, Dict[str, str]]:
    store = _load_json(_TRAIN_FILE)
    rows = store.get("rows", []) if isinstance(store, dict) else []
    result: Dict[str, Dict[str, str]] = {}
    counts: Dict[str, Dict[str, Dict[str, int]]] = {}
    for r in rows:
        tid = r.get("template_id")
        f = r.get("field")
        v = r.get("value")
        if not tid or not f or not v:
            continue
        counts.setdefault(tid, {}).setdefault(f, {}).setdefault(v, 0)
        counts[tid][f][v] += 1
    for tid, fd in counts.items():
        result.setdefault(tid, {})
        for f, cdict in fd.items():
            best = None
            freq = -1
            for v, n in cdict.items():
                if n > freq:
                    best = v
                    freq = n
            if best:
                result[tid][f] = best
    _save_json(_LEARNED_FILE, result)
    return result


def _rand_name() -> str:
    xs = ["张三", "李四", "王五", "赵六", "孙七", "周八", "吴九", "郑十", "钱一", "刘二"]
    return random.choice(xs)


def _rand_org() -> str:
    xs = ["××大学", "××公司", "××研究院", "××实验室"]
    return random.choice(xs)


def _rand_text(kind: str) -> str:
    base = {
        "案由": ["合同纠纷", "劳动争议", "侵权纠纷"],
        "诉讼请求": ["请求承担损失", "请求返还款项", "请求解除合同"],
        "请假类型": ["事假", "病假", "年休假"],
        "部门": ["研发部", "市场部", "人事部"],
        "会议地点": ["会议室A", "会议室B", "线上会议"],
        "主持人": ["主持人甲", "主持人乙"],
        "学校与专业": ["××大学计算机", "××学院数据科学", "××大学电子信息"],
        "实习岗位": ["数据分析", "算法工程", "前端开发"],
        "实习单位": ["××科技", "××互联网", "××制造"],
        "经费预算": ["5万", "10万", "20万"],
        "评估指标": ["MAE", "RMSE", "F1"],
    }
    xs = base.get(kind)
    if xs:
        return random.choice(xs)
    return f"示例{kind}"


def synthesize_training_data(per_template: int = 20) -> int:
    cnt = 0
    for t in list_templates():
        for _ in range(per_template):
            data: Dict[str, str] = {}
            for f in t.fields:
                if f in {"原告姓名", "被告姓名", "委托人姓名", "受托人姓名", "申请人姓名", "负责人", "推荐人姓名", "被推荐人姓名"}:
                    data[f] = _rand_name()
                elif f in {"原告性别"}:
                    data[f] = random.choice(["男", "女"])
                elif f in {"甲方名称", "乙方名称", "推荐人单位"}:
                    data[f] = _rand_org()
                elif f in {"法院名称"}:
                    data[f] = "××人民法院"
                elif f in {"日期", "申请日期"}:
                    data[f] = _today()
                else:
                    data[f] = _rand_text(f)
            record_training(t.id, data)
            cnt += 1
    return cnt


def auto_train(per_template: int = 20) -> Dict[str, Dict[str, str]]:
    synthesize_training_data(per_template)
    return run_training()


_HW_STYLE_FILE = "handwrite_style.json"
_HISTORY_FILE = "history.json"


def _resource_base() -> str:
    p = getattr(sys, "_MEIPASS", None)
    if p and os.path.isdir(p):
        return p
    return os.path.dirname(os.path.abspath(__file__))


def _fonts_dir() -> str:
    base = _resource_base()
    packaged = os.path.join(base, "assets", "fonts")
    if os.path.isdir(packaged):
        return packaged
    writable = os.path.join(os.getcwd(), "assets", "fonts")
    os.makedirs(writable, exist_ok=True)
    return writable


def _font_urls() -> List[Dict[str, str]]:
    return [
        {
            "name": "MaShanZheng-Regular.ttf",
            "url": "https://github.com/google/fonts/raw/main/ofl/mashanzheng/MaShanZheng-Regular.ttf",
        },
        {
            "name": "ZhiMangXing-Regular.ttf",
            "url": "https://github.com/google/fonts/raw/main/ofl/zhimangxing/ZhiMangXing-Regular.ttf",
        },
        {
            "name": "LongCang-Regular.ttf",
            "url": "https://github.com/google/fonts/raw/main/ofl/longcang/LongCang-Regular.ttf",
        },
    ]


def ensure_handwrite_assets() -> List[str]:
    dirp = _fonts_dir()
    paths: List[str] = []
    # Prefer packaged fonts if present
    for it in _font_urls():
        p = os.path.join(dirp, it["name"])
        if os.path.exists(p):
            paths.append(p)
    # If none found, attempt download into writable dir
    if not paths:
        for it in _font_urls():
            p = os.path.join(dirp, it["name"])
            if not os.path.exists(p):
                try:
                    urllib.request.urlretrieve(it["url"], p)
                except Exception:
                    pass
            if os.path.exists(p):
                paths.append(p)
    return paths

def _win_fonts_dir() -> str:
    return os.path.join(os.environ.get("WINDIR", "C:\\Windows"), "Fonts")

def _mac_fonts_dirs() -> List[str]:
    homes = os.path.expanduser("~")
    return [
        "/System/Library/Fonts",
        "/Library/Fonts",
        os.path.join(homes, "Library", "Fonts"),
    ]

def resolve_font_by_name(name: str) -> Optional[str]:
    nm = (name or "").strip().lower()
    d = _win_fonts_dir()
    if nm in {"宋体", "song", "simsun"}:
        p = os.path.join(d, "simsun.ttc")
        if os.path.exists(p):
            return p
        # Fallback to SimHei if SimSun TTC not available
        p2 = os.path.join(d, "simhei.ttf")
        return p2 if os.path.exists(p2) else None
    if nm in {"楷体", "kaiti", "simkai"}:
        p = os.path.join(d, "simkai.ttf")
        return p if os.path.exists(p) else None
    if nm in {"黑体", "hei", "simhei"}:
        p = os.path.join(d, "simhei.ttf")
        return p if os.path.exists(p) else None
    if nm in {"手写-马善政"}:
        p = os.path.join(_fonts_dir(), "MaShanZheng-Regular.ttf")
        return p if os.path.exists(p) else None
    if nm in {"手写-芝蔓行"}:
        p = os.path.join(_fonts_dir(), "ZhiMangXing-Regular.ttf")
        return p if os.path.exists(p) else None
    if nm in {"手写-龙藏"}:
        p = os.path.join(_fonts_dir(), "LongCang-Regular.ttf")
        return p if os.path.exists(p) else None
    # macOS candidates
    if sys.platform == "darwin":
        cands: List[str] = []
        if nm in {"宋体", "song", "simsun"}:
            cands = [
                "Songti.ttc",
                "STSong.ttf",
                "STSongti-SC-Regular.otf",
                "PingFang.ttc",
            ]
        elif nm in {"楷体", "kaiti", "simkai"}:
            cands = [
                "Kaiti.ttc",
                "STKaiti.ttf",
                "STKaiti-SC-Regular.otf",
            ]
        elif nm in {"黑体", "hei", "simhei"}:
            cands = [
                "STHeiti Light.ttc",
                "PingFang.ttc",
                "Heiti.ttc",
            ]
        for base in _mac_fonts_dirs():
            for fn in cands:
                p = os.path.join(base, fn)
                if os.path.exists(p):
                    return p
    return None


def train_handwrite_style() -> Dict[str, str]:
    fonts = ensure_handwrite_assets()
    if not fonts:
        raise RuntimeError("未能下载手写体字体资源")
    seed = int(time.time())
    random.seed(seed)
    style = {
        "font": random.choice(fonts),
        "font_size": random.choice([36, 40, 44]),
        "line_gap": random.choice([14, 18, 22]),
        "rotate_min": -2,
        "rotate_max": 2,
        "jitter": random.choice([0, 1, 2]),
    }
    _save_json(_HW_STYLE_FILE, style)
    return style


def _load_handwrite_style() -> Dict[str, str]:
    cfg = _load_json(_HW_STYLE_FILE)
    if not isinstance(cfg, dict) or not cfg.get("font"):
        cfg = train_handwrite_style()
    return cfg


def generate_handwriting_image(text: str, out_path: str, style: Optional[Dict[str, str]] = None) -> str:
    try:
        from PIL import Image, ImageDraw, ImageFont
    except Exception:
        raise RuntimeError("未检测到Pillow，请先安装：pip install pillow")
    cfg = style or _load_handwrite_style()
    font_path = cfg.get("font")
    font_size = int(cfg.get("font_size", 40))
    line_gap = int(cfg.get("line_gap", 18))
    rmin = float(cfg.get("rotate_min", -1))
    rmax = float(cfg.get("rotate_max", 1))
    jitter = int(cfg.get("jitter", 1))
    if not font_path or not os.path.exists(font_path):
        fname = (style or {}).get("font_name") if style else None
        if fname:
            fp = resolve_font_by_name(fname)
            if fp and os.path.exists(fp):
                font_path = fp
    
    if not font_path or not os.path.exists(font_path):
        fonts = ensure_handwrite_assets()
        if not fonts:
            raise RuntimeError("无可用手写体字体")
        font_path = fonts[0]
    # Pillow may have limited support for TTC; prefer TTF
    if font_path.lower().endswith(".ttc"):
        alt = os.path.join(_win_fonts_dir(), "simhei.ttf")
        if os.path.exists(alt):
            font_path = alt
    font = ImageFont.truetype(font_path, font_size)
    lines = [x for x in (text or "").splitlines() if x.strip()]
    if not lines:
        lines = [" "]
    max_chars = max(len(x) for x in lines)
    w = max(800, int(max_chars * font_size * 0.7))
    h = int(len(lines) * (font_size + line_gap) + 40)
    img = Image.new("RGB", (w, h), color=(255, 255, 255))
    draw = ImageDraw.Draw(img)
    y = 20
    for i, line in enumerate(lines):
        dy = y + random.randint(-jitter, jitter)
        rot = random.uniform(rmin, rmax)
        tmp = Image.new("RGBA", (w, font_size + 8), (255, 255, 255, 0))
        tdraw = ImageDraw.Draw(tmp)
        tdraw.text((20 + random.randint(0, jitter), 0), line, font=font, fill=(0, 0, 0))
        tmp = tmp.rotate(rot, resample=Image.BICUBIC, expand=1)
        img.paste(Image.alpha_composite(Image.new("RGBA", tmp.size, (255, 255, 255, 0)), tmp), (0, dy), tmp)
        y += font_size + line_gap
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    img.save(out_path)
    return out_path


def auto_generate_image_for_document(template_id: str, data: Dict[str, str], style: str = "formal", out_dir: Optional[str] = None) -> str:
    txt = generate_document(template_id, data, style)
    d = out_dir or os.path.join(os.getcwd(), "handwrite_output")
    os.makedirs(d, exist_ok=True)
    fname = f"{template_id}_{int(time.time())}.png"
    path = os.path.join(d, fname)
    return generate_handwriting_image(txt, path)


def batch_generate_images(input_dir: str, output_dir: str, encoding: str = "utf-8") -> List[str]:
    outs: List[str] = []
    if not os.path.isdir(input_dir):
        return outs
    os.makedirs(output_dir, exist_ok=True)
    for root, _, files in os.walk(input_dir):
        for fn in files:
            if fn.lower().endswith(".txt"):
                ip = os.path.join(root, fn)
                try:
                    with open(ip, "r", encoding=encoding) as f:
                        txt = f.read()
                    op = os.path.join(output_dir, os.path.splitext(fn)[0] + ".png")
                    outs.append(generate_handwriting_image(txt, op))
                except Exception:
                    pass
    return outs

def add_history(template_id: str, data: Dict[str, str], text: str, image_path: Optional[str]) -> None:
    item = {
        "ts": int(time.time()),
        "template_id": template_id,
        "data": data,
        "text": text,
        "image_path": image_path or "",
    }
    store = _load_json(_HISTORY_FILE)
    if not isinstance(store, dict):
        store = {}
    lst = store.get("items", [])
    if not isinstance(lst, list):
        lst = []
    lst.append(item)
    lst = lst[-100:]
    store["items"] = lst
    _save_json(_HISTORY_FILE, store)

def list_history() -> List[Dict[str, str]]:
    store = _load_json(_HISTORY_FILE)
    items = store.get("items", []) if isinstance(store, dict) else []
    items = sorted(items, key=lambda x: x.get("ts", 0), reverse=True)
    return items

def latest_history_for_template(template_id: str) -> Optional[Dict[str, str]]:
    for it in list_history():
        if it.get("template_id") == template_id:
            return it
    return None

