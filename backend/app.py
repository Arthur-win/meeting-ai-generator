from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import subprocess
import json
import re
import os
import time
from word_generator import WordMeetingGenerator

app = Flask(__name__)
CORS(app)  # 启用CORS支持

def call_ollama(prompt, model="llama3"):
    """调用本地 Ollama 模型"""
    try:
        result = subprocess.run(
            ["ollama", "run", model],
            input=prompt.encode("utf-8"),
            capture_output=True,
            timeout=120  # 最长等待 120 秒
        )
        output = result.stdout.decode("utf-8", errors="ignore")
        # 安全地清理输出，移除可能的模型前缀信息
        try:
            # 避免空输出或无效正则表达式导致的错误
            if output:
                output = re.sub(r'^.*?---', '', output, flags=re.DOTALL)
        except re.error as e:
            print(f"清理输出时的正则表达式错误: {e}")
        return output.strip() if output else ""
    except subprocess.TimeoutExpired:
        return "模型响应超时，请检查 Ollama 是否正常运行。"
    except Exception as e:
        print(f"调用Ollama详细错误: {e}")
        return f"调用 Ollama 出错：{str(e)}"

def parse_meeting_info(meeting_text):
    """解析会议信息（支持多种自然语言描述方式和不同会议类型）"""
    # 初始化会议信息字典，设置默认值
    meeting_info = {
        'theme': '未指定',
        'host': '未指定', 
        'location': '未指定',
        'attendees': '未指定',
        'duration': '未指定',
        'topics': [],
        'meeting_type': '通用会议'  # 新增会议类型字段
    }
    
    # 确保会议文本是有效的字符串
    if not meeting_text or not isinstance(meeting_text, str):
        print("无效的会议文本")
        return meeting_info
    
    # 打印解析文本以调试
    print("解析文本:", repr(meeting_text))
    
    # 安全的正则表达式搜索函数
    def safe_re_search(pattern, text, flags=0):
        try:
            return re.search(pattern, text, flags)
        except re.error as e:
            print(f"正则表达式错误: {pattern}, {e}")
            return None
    
    # 安全的正则表达式替换函数
    def safe_re_sub(pattern, repl, text, flags=0):
        try:
            return re.sub(pattern, repl, text, flags)
        except re.error as e:
            print(f"正则表达式替换错误: {pattern}, {repl}, {e}")
            return text
    
    # 安全的正则表达式分割函数
    def safe_re_split(pattern, text, maxsplit=0, flags=0):
        try:
            return re.split(pattern, text, maxsplit=maxsplit, flags=flags)
        except re.error as e:
            print(f"正则表达式分割错误: {pattern}, {e}")
            return [text]
    
    # 安全的正则表达式查找所有函数
    def safe_re_findall(pattern, text, flags=0):
        try:
            return re.findall(pattern, text, flags=flags)
        except re.error as e:
            print(f"正则表达式查找错误: {pattern}, {e}")
            return []
    
    # ===== 检测会议类型 =====
    # 技术会议关键词
    tech_keywords = ['技术', '开发', '编程', '代码', 'API', '架构', '数据库', '前端', '后端', '测试', 
                     'bug', '修复', '部署', '性能', '优化', '迭代', '版本', '需求', '设计']
    # 商务会议关键词
    business_keywords = ['商务', '合作', '谈判', '客户', '市场', '销售', '营销', '推广', '策略', '预算',
                        '财务', '投资', '盈利', '成本', '分析', '报告', '季度', '年度', '计划']
    # 项目会议关键词
    project_keywords = ['项目', '进度', '里程碑', '任务', '分工', '责任', '延期', '风险', '协调', 
                       '资源', '分配', '时间线', '甘特图', '交付', '验收']
    # 团队会议关键词
    team_keywords = ['团队', '部门', '周会', '例会', '分享', '讨论', '交流', '培训', '总结',
                    '回顾', '展望', '问题', '建议', '反馈']
    
    # 检测会议类型
    keyword_count = {
        '技术会议': sum(keyword in meeting_text for keyword in tech_keywords),
        '商务会议': sum(keyword in meeting_text for keyword in business_keywords),
        '项目会议': sum(keyword in meeting_text for keyword in project_keywords),
        '团队会议': sum(keyword in meeting_text for keyword in team_keywords)
    }
    
    # 确定主要会议类型
    if max(keyword_count.values()) > 1:  # 需要至少有2个关键词匹配才算特定类型
        meeting_info['meeting_type'] = max(keyword_count.items(), key=lambda x: x[1])[0]
    
    print(f"检测到会议类型: {meeting_info['meeting_type']}")
    
    # ===== 提取主持人 =====
    # 自然语言中的主持人模式
    host_patterns = [
        r'主持人[：:]\s*(.+?)[\n\r]',
        r'主持人[：:]\s*(.+)',
        r'主持[：:]\s*(.+?)[\n\r]',
        r'主持[：:]\s*(.+)',
        r'由([^，,。\n\r]+)主持',
        r'([^，,。\n\r]+)主持会议',
        r'会议主持[：:]\s*(.+?)[\n\r]',
        r'会议主持[：:]\s*(.+)'  
    ]
    
    for pattern in host_patterns:
        match = safe_re_search(pattern, meeting_text)
        if match:
            meeting_info['host'] = match.group(1).strip()
            print(f"提取到主持人: {meeting_info['host']}")
            break
    
    # ===== 提取会议地点 =====
    # 自然语言中的会议地点模式
    location_patterns = [
        # 优先匹配"在...召开"格式，解决'公司三楼大会议室'问题
        r'(?:在|于)\s*([^，,。\n\r]+?)\s*(?:召开|举行|组织|进行|开展)',
        # 匹配"在...会议室"格式
        r'在([^，,。\n\r]+会议室)',
        r'在([^，,。\n\r]*?(?:会议室|多功能厅|办公室|厅|楼|室|中心|院|馆|房|场|所|部))',
        # 基础地点提取模式
        r'会议地点[：:]\s*(.+?)[\n\r]',
        r'会议地点[：:]\s*(.+)',
        r'地点[：:]\s*(.+?)[\n\r]',
        r'地点[：:]\s*(.+)',
        r'会议将在([^，,。\n\r]+)举行',
        r'会议场地[：:]\s*(.+?)[\n\r]',
        r'会议场地[：:]\s*(.+)',
        r'在([^，,。\n\r]+)进行',
        r'聚会地点[：:]\s*(.+?)[\n\r]',
        r'聚会地点[：:]\s*(.+)',
        # 提取"时间，地点"格式
        r'[点分][，,]\s*(?:在)?([^，,。\n\r]+?)(?:举行|召开)',
        # 兜底模式：匹配"在...召开"的变体
        r'在([^，,。\n\r]+)召开',
        r'在([^，,。\n\r]+)举行'
    ]
    
    for pattern in location_patterns:
        match = safe_re_search(pattern, meeting_text)
        if match:
            meeting_info['location'] = match.group(1).strip()
            print(f"提取到地点: {meeting_info['location']}")
            break
    
    # ===== 提取参会人员 =====
    # 自然语言中的参会人员模式
    attendees_patterns = [
        # 基础参会人员提取模式
        r'参会人员[：:]\s*([^，,。\n\r]+)',
        r'参会的有([^，,。\n\r]+)',
        r'参加的有([^，,。\n\r]+)',
        r'出席人员[：:]\s*([^，,。\n\r]+)',
        r'参与人员[：:]\s*([^，,。\n\r]+)',
        r'参加人员包括([^，,。\n\r]+)',
        r'到场的有([^，,。\n\r]+)',
        r'参与的有([^，,。\n\r]+)',
        r'出席的有([^，,。\n\r]+)',
        r'参会者[：:]\s*([^，,。\n\r]+)',
        # 增强模式：部门分组和多人名单
        r'(?:参会|参加|出席|与会)(?:人员)?[:：]\s*([^。\n\r]+)',
        r'(?:参会|参加|出席|与会)(?:的)?(?:人|人员)?(?:为|是|包括)?\s*([^。\n\r]+)',
        # 提取包含部门信息的参会人员格式
        r'(?:[，,]\s*|[。]\s*)([^，,。\n\r]+?)(?:的|部门的)([^，,。\n\r]+?)(?:[，,。]|以及)',
        # 提取"以及"后面的人员
        r'以及\s*([^，,。\n\r]+)',
        # 新增更精确的模式，匹配"参会人员有X、Y、Z"格式
        r'参会人员有([^，,。\n\r]+?)(?:，会议大概开|，会议|。)',
        # 新增更精确的模式，匹配"参会人员：X、Y、Z"格式
        r'参会人员[：:]\s*([^，,。\n\r]+?)(?:，会议大概开|，会议|。)'
    ]
    
    # 收集所有可能的参会人员信息
    all_attendees = []
    
    # 首先尝试提取主要的参会人员信息
    for pattern in attendees_patterns:
        match = safe_re_search(pattern, meeting_text)
        if match:
            if len(match.groups()) == 2:
                # 处理包含部门信息的匹配结果
                department = match.group(1).strip()
                people = match.group(2).strip()
                # 分割多个人员
                for person in re.split(r'[，,]', people):
                    person = person.strip()
                    if person and not any(x in person for x in ['等', '以及', '会议大概开', '小时']):
                        all_attendees.append(f"{department}的{person}")
            else:
                attendees_text = match.group(1).strip()
                # 分割逗号分隔的名单
                for person in safe_re_split(r'[，,]', attendees_text):
                    person = person.strip()
                    # 过滤掉不相关的文本，如"会议大概开两个小时"
                    if person and not any(x in person for x in ['等', '以及', '会议大概开', '小时', '会议', '大概', '开']):
                        # 清理参会人员文本，移除不必要的修饰词
                        person = safe_re_sub(r'[的是还有等]+', '', person)
                        # 处理部门前缀，如"市场部李明" -> "市场部的李明"
                        if not person.startswith('市场部的') and '市场部' in person and not '的' in person:
                            person = person.replace('市场部', '市场部的')
                        # 清理掉可能的"员"字前缀
                        if person.startswith('员'):
                            person = person[1:]
                        all_attendees.append(person)
    
    # 检查"以及"后面的人员（额外处理）
    try:
        and_pattern = r'以及\s*([^，,。\n\r]+)'
        and_matches = safe_re_findall(and_pattern, meeting_text, re.MULTILINE)
        for match in and_matches:
            for person in safe_re_split(r'[，,]', match.strip()):
                person = person.strip()
                if person and not any(x in person for x in ['等', '以及']):
                    all_attendees.append(person)
    except re.error as e:
        print(f"参会人员'以及'正则表达式错误: {e}")
    
    # 如果有收集到参会人员，去重并合并为字符串
    if all_attendees:
        # 去重并合并为字符串
        unique_attendees = list(dict.fromkeys(all_attendees))  # 保持顺序的去重
        
        # 进一步清理重复的人员名称（如"李明"和"市场部的李明"）
        cleaned_attendees = []
        for attendee in unique_attendees:
            # 检查是否已经有更简单的版本存在
            simple_name = attendee.split('的')[-1] if '的' in attendee else attendee
            # 检查是否已经有这个简单名称或更详细的版本
            if not any(simple_name in existing or existing in attendee for existing in cleaned_attendees):
                cleaned_attendees.append(attendee)
        
        meeting_info['attendees'] = '，'.join(cleaned_attendees)
        print(f"提取到参会人员: {meeting_info['attendees']}")
    else:
        # 如果没有找到任何匹配，继续使用默认值
        print("未提取到明确的参会人员信息")
    
    # ===== 提取会议时长 =====
    # 自然语言中的会议时长模式
    duration_patterns = [
        r'会议大概开([^，,。\n\r]+)',
        r'会议时长[：:]\s*(.+?)[\n\r]',
        r'会议时长[：:]\s*(.+)',
        r'时长[：:]\s*(.+?)[\n\r]',
        r'时长[：:]\s*(.+)',
        r'大约([^，,。\n\r]+)小时',
        r'预计([^，,。\n\r]+)',
        r'会议持续([^，,。\n\r]+)',
        r'持续([^，,。\n\r]+)',
        r'大概开([^，,。\n\r]+)',
        r'会议预计([^，,。\n\r]+)',
        r'将持续([^，,。\n\r]+)'
    ]
    
    for pattern in duration_patterns:
        match = safe_re_search(pattern, meeting_text)
        if match:
            meeting_info['duration'] = match.group(1).strip()
            print(f"提取到时长: {meeting_info['duration']}")
            break
    
    # ===== 提取会议主题 =====
    # 如果没有明确的主题，尝试从议题中推断
    theme_patterns = [
        # 基础会议主题提取模式
        r'会议主题[：:]\s*(.+?)[\n\r]',
        r'会议主题[：:]\s*(.+)',
        r'主题[：:]\s*(.+?)[\n\r]',
        r'主题[：:]\s*(.+)',
        r'讨论关于([^，,。\n\r]+)的会议',
        r'讨论([^，,。\n\r]+)',
        r'关于([^，,。\n\r]+)的会议',
        r'议题是([^，,。\n\r]+)',
        r'议题为([^，,。\n\r]+)',
        r'会议议题[：:]\s*(.+?)[\n\r]',
        r'会议议题[：:]\s*(.+)',
        r'主要讨论([^，,。\n\r]+)',
        r'重点讨论([^，,。\n\r]+)',
        # 增强会议主题识别
        r'(?:召开|举行|组织|进行|开展|开)(?:\s*|的)?(?:一次|一个|一场)?\s*([^，,。\n\r]+?)(?:会议|会)(?!\s*内容)',
        r'(?:召开|举行|组织|进行|开展|开)(?:\s*|的)?(?:一次|一个|一场)?\s*([^，,。\n\r]+?)[，,。\n\r]',
        r'(?:关于|就)\s*([^的的会议]+?)\s*的会议',
        r'在[^，,。\n\r]+?召开(?:一次|一个|一场)?\s*([^，,。\n\r]+?)(?:会议|会)(?!\s*内容)',
        # 其他常见会议主题模式
        r'召开的(?:是)?\s*([^，,。\n\r]+?)会议',
        r'举行的(?:是)?\s*([^，,。\n\r]+?)会议',
        r'为了([^，,。\n\r]+?)召开会议',
        r'为了([^，,。\n\r]+?)举行会议',
        r'基于([^，,。\n\r]+?)召开会议',
        r'围绕([^，,。\n\r]+?)展开讨论',
        r'聚焦([^，,。\n\r]+?)(?:的讨论)?',
        r'会议围绕([^，,。\n\r]+?)进行'
    ]
    
    # 针对不同会议类型的特定主题提取模式
    if meeting_info['meeting_type'] == '技术会议':
        theme_patterns.extend([
            r'技术讨论[：:]\s*(.+?)[\n\r]',
            r'技术讨论[：:]\s*(.+)',
            r'开发([^，,。\n\r]+)',
            r'需求评审[：:]\s*(.+?)[\n\r]',
            r'需求评审[：:]\s*(.+)',
            r'架构设计[：:]\s*(.+?)[\n\r]',
            r'架构设计[：:]\s*(.+)'
        ])
    elif meeting_info['meeting_type'] == '商务会议':
        theme_patterns.extend([
            r'商务洽谈[：:]\s*(.+?)[\n\r]',
            r'商务洽谈[：:]\s*(.+)',
            r'合作([^，,。\n\r]+)',
            r'市场分析[：:]\s*(.+?)[\n\r]',
            r'市场分析[：:]\s*(.+)',
            r'项目报价[：:]\s*(.+?)[\n\r]',
            r'项目报价[：:]\s*(.+)'
        ])
    elif meeting_info['meeting_type'] == '项目会议':
        theme_patterns.extend([
            r'项目进度[：:]\s*(.+?)[\n\r]',
            r'项目进度[：:]\s*(.+)',
            r'项目评审[：:]\s*(.+?)[\n\r]',
            r'项目评审[：:]\s*(.+)',
            r'里程碑([^，,。\n\r]+)',
            r'项目计划[：:]\s*(.+?)[\n\r]',
            r'项目计划[：:]\s*(.+)'
        ])
    elif meeting_info['meeting_type'] == '团队会议':
        theme_patterns.extend([
            r'周会[：:]\s*(.+?)[\n\r]',
            r'周会[：:]\s*(.+)',
            r'部门会议[：:]\s*(.+?)[\n\r]',
            r'部门会议[：:]\s*(.+)',
            r'团队分享[：:]\s*(.+?)[\n\r]',
            r'团队分享[：:]\s*(.+)'
        ])
    
    for pattern in theme_patterns:
        match = safe_re_search(pattern, meeting_text)
        if match:
            meeting_info['theme'] = match.group(1).strip()
            print(f"提取到主题: {meeting_info['theme']}")
            break
    
    # ===== 提取议题 =====
    topics = []
    
    # 1. 尝试匹配"会议主要说X件事"或"主要讨论X点"后面的议题
    main_topics_match = safe_re_search(r'(会议主要说|主要讨论)[^，,。\n\r]*[事点]，([^。]+)', meeting_text)
    if main_topics_match:
        topics_part = main_topics_match.group(2).strip()
        # 使用中文数字（一、二、三）分割议题
        try:
            topic_segments = safe_re_split(r'[一二三四五六七八九十][是、.]', topics_part)
            topic_segments = [t.strip() for t in topic_segments if t.strip()]
        except re.error as e:
            print(f"议题分割正则表达式错误: {e}")
            topic_segments = []
        
        for i, topic_segment in enumerate(topic_segments, 1):
            # 提取议题内容
            topic_text_match = re.search(r'^(.+?)(，|；|。|$)', topic_segment)
            if topic_text_match:
                topic_text = topic_text_match.group(1).strip()
                
                # 创建议题对象
                topic_obj = {
                    'topic': topic_text,
                    'leader': '未指定',
                    'preparation': '无'
                }
                topics.append(topic_obj)
                print(f"提取到议题{i}: {topic_text}")
    
    # 2. 如果没有提取到议题，尝试其他格式
    if not topics:
        # 尝试直接匹配数字编号的议题
        for num in range(1, 11):  # 最多10个议题
            # 尝试多种格式的正则表达式
            patterns = [
                rf'议题{num}[、.。:：]\s*(.+?)[\n\r]',
                rf'议题{num}[、.。:：]\s*(.+)',
                rf'{num}[是、.]\s*(.+?)[，,；；。。\n\r]',
                rf'{num}[是、.]\s*(.+)'  
            ]
            
            for pattern in patterns:
                match = safe_re_search(pattern, meeting_text)
                if match:
                    topic_text = match.group(1).strip()
                    topic_obj = {
                        'topic': topic_text,
                        'leader': '未指定',
                        'preparation': '无'
                    }
                    topics.append(topic_obj)
                    print(f"提取到议题{num}: {topic_text}")
                    break
    
    # 3. 尝试匹配中文数字格式的议题
    if not topics:
        for i, chinese_num in enumerate(['一', '二', '三', '四', '五', '六', '七', '八', '九', '十'], 1):
            try:
                patterns = [
                    rf'{chinese_num}[、.]\s*(.+?)[，,；；。。\n\r]',
                    rf'{chinese_num}[、.]\s*(.+)',
                    rf'{chinese_num}是\s*(.+?)[，,；；。。\n\r]',
                    rf'{chinese_num}是\s*(.+)',
                    rf'{chinese_num}、([^，,；；。。\n\r]+)',
                    rf'第{chinese_num}[项条].*?[：:]([^，,；；。。\n\r]+)'
                ]
                
                # 针对不同会议类型的特定议题提取模式
                if meeting_info['meeting_type'] == '技术会议':
                    patterns.extend([
                        rf'{chinese_num}[、.]\s*技术([^，,；；。。\n\r]+)',
                        rf'{chinese_num}[、.]\s*开发([^，,；；。。\n\r]+)',
                        rf'{chinese_num}[、.]\s*问题([^，,；；。。\n\r]+)',
                        rf'{chinese_num}[、.]\s*修复([^，,；；。。\n\r]+)'
                    ])
                elif meeting_info['meeting_type'] == '商务会议':
                    patterns.extend([
                        rf'{chinese_num}[、.]\s*商务([^，,；；。。\n\r]+)',
                        rf'{chinese_num}[、.]\s*合作([^，,；；。。\n\r]+)',
                        rf'{chinese_num}[、.]\s*客户([^，,；；。。\n\r]+)',
                        rf'{chinese_num}[、.]\s*市场([^，,；；。。\n\r]+)'
                    ])
                elif meeting_info['meeting_type'] == '项目会议':
                    patterns.extend([
                        rf'{chinese_num}[、.]\s*进度([^，,；；。。\n\r]+)',
                        rf'{chinese_num}[、.]\s*任务([^，,；；。。\n\r]+)',
                        rf'{chinese_num}[、.]\s*问题([^，,；；。。\n\r]+)',
                        rf'{chinese_num}[、.]\s*资源([^，,；；。。\n\r]+)'
                    ])
                elif meeting_info['meeting_type'] == '团队会议':
                    patterns.extend([
                        rf'{chinese_num}[、.]\s*分享([^，,；；。。\n\r]+)',
                        rf'{chinese_num}[、.]\s*问题([^，,；；。。\n\r]+)',
                        rf'{chinese_num}[、.]\s*计划([^，,；；。。\n\r]+)',
                        rf'{chinese_num}[、.]\s*总结([^，,；；。。\n\r]+)'
                    ])
                
                for pattern in patterns:
                    match = safe_re_search(pattern, meeting_text)
                    if match:
                        topic_text = match.group(1).strip()
                        if topic_text.strip():
                            topic_obj = {
                                'topic': topic_text,
                                'leader': '未指定',
                                'preparation': '无'
                            }
                            topics.append(topic_obj)
                        print(f"提取到议题{i}: {topic_text}")
                        break
            except Exception as e:
                print(f"中文数字议题提取错误: {e}")
                continue
    
    # ===== 为每个议题提取负责人 =====
    for i, topic in enumerate(topics):
        # 增强的负责人提取模式列表，确保议题编号与负责人正确关联
        leader_patterns = [
            # 直接关联议题编号和负责人的模式（优先匹配）
            rf'议题{i+1}[、.。:：]?\s*.*?负责人[：:]\s*([^，,。\n\r]+)',
            rf'议题{i+1}[、.。:：]?\s*.*?由([^，,。\n\r]+)负责',
            rf'{i+1}[、.]\s*.*?由([^，,。\n\r]+)负责',
            rf'{i+1}[是、.]\s*.+?，([^，,。\n\r]+)负责',
            rf'{i+1}[是、.]\s*.+?，([^，,。\n\r]+)[你您]负责',
            rf'{i+1}[是、.]\s*.+?，([^，,。\n\r]+)[你您]牵头',
            rf'{re.escape(topic["topic"])}.*?，([^，,。\n\r]+)负责',
            rf'{re.escape(topic["topic"])}.*?，([^，,。\n\r]+)[你您]负责',
            rf'{re.escape(topic["topic"])}.*?，([^，,。\n\r]+)[你您]牵头',
            rf'负责([^，,。\n\r]+)，{re.escape(topic["topic"])}',
            rf'由([^，,。\n\r]+)负责.*?{re.escape(topic["topic"])}',
            rf'{re.escape(topic["topic"])}.*?由([^，,。\n\r]+)负责',
            rf'{re.escape(topic["topic"])}.*?由([^，,。\n\r]+)[你您]负责',
            rf'{re.escape(topic["topic"])}.*?由([^，,。\n\r]+)[你您]牵头',
            rf'{re.escape(topic["topic"])}.*?([^，,。\n\r]+)负责',
            rf'{re.escape(topic["topic"])}.*?([^，,。\n\r]+)牵头',
            # 增加更精确的议题编号与负责人关联模式
            rf'第{i+1}[项个点].*?负责人[：:]\s*([^，,。\n\r]+)',
            rf'第{i+1}[项个点].*?由([^，,。\n\r]+)负责',
            # 增加针对常见负责人描述的模式
            rf'{i+1}[、.]\s*.*?([^，,。\n\r]+)(?:负责|牵头|主讲|汇报)该议题',
            rf'议题{i+1}\s*([^，,。\n\r]+)(?:负责|牵头|主讲|汇报)'
        ]
        
        # 针对不同会议类型的特定负责人提取模式
        if meeting_info['meeting_type'] == '技术会议':
            leader_patterns.extend([
                rf'技术负责人[：:]\s*(.+?)[\n\r]',
                rf'技术负责人[：:]\s*(.+)',
                rf'开发负责人[：:]\s*(.+?)[\n\r]',
                rf'开发负责人[：:]\s*(.+)',
                rf'工程师[：:]\s*(.+?)[\n\r]',
                rf'工程师[：:]\s*(.+)'
            ])
        elif meeting_info['meeting_type'] == '商务会议':
            leader_patterns.extend([
                rf'商务负责人[：:]\s*(.+?)[\n\r]',
                rf'商务负责人[：:]\s*(.+)',
                rf'客户经理[：:]\s*(.+?)[\n\r]',
                rf'客户经理[：:]\s*(.+)',
                rf'销售负责人[：:]\s*(.+?)[\n\r]',
                rf'销售负责人[：:]\s*(.+)'  
            ])
        elif meeting_info['meeting_type'] == '项目会议':
            leader_patterns.extend([
                rf'项目经理[：:]\s*(.+?)[\n\r]',
                rf'项目经理[：:]\s*(.+)',
                rf'项目负责人[：:]\s*(.+?)[\n\r]',
                rf'项目负责人[：:]\s*(.+)',
                rf'负责人[：:]\s*(.+?)[\n\r]',
                rf'负责人[：:]\s*(.+)'  
            ])
        elif meeting_info['meeting_type'] == '团队会议':
            leader_patterns.extend([
                rf'团队负责人[：:]\s*(.+?)[\n\r]',
                rf'团队负责人[：:]\s*(.+)',
                rf'部门负责人[：:]\s*(.+?)[\n\r]',
                rf'部门负责人[：:]\s*(.+)',
                rf'组长[：:]\s*(.+?)[\n\r]',
                rf'组长[：:]\s*(.+)'  
            ])
        
        for pattern in leader_patterns:
            match = safe_re_search(pattern, meeting_text)
            if match:
                leader = match.group(1).strip()
                # 清理负责人名称，移除"负责"、"牵头"等词语
                leader = safe_re_sub(r'[你您]', '', leader)
                leader = safe_re_sub(r'负责$', '', leader)
                leader = safe_re_sub(r'牵头$', '', leader)
                leader = safe_re_sub(r'主讲$', '', leader)
                leader = safe_re_sub(r'汇报$', '', leader)
                leader = safe_re_sub(r'^由', '', leader)
                leader = leader.strip()
                topic['leader'] = leader if leader else '未指定'
                print(f"提取到负责人: {leader}")
                break
    
    # ===== 提取全局准备事项（适用于所有议题） =====
    global_preparations = []
    global_prep_patterns = [
        r'(?:请|请务必|请大家)?\s*提前将([^，,。\n\r]+)',
        r'(?:请|请务必|请大家)?\s*提前准备([^，,。\n\r]+)',
        r'会前准备[：:]\s*([^，,。\n\r]+)',
        r'所有参会人员.*?准备([^，,。\n\r]+)',
        r'会议准备[：:]\s*([^，,。\n\r]+)',
        r'(?:总结报告|规划草案|材料|文档)发至([^，,。\n\r]+)',
    ]
    
    for pattern in global_prep_patterns:
        match = safe_re_search(pattern, meeting_text)
        if match:
            global_preparations.append(match.group(1).strip())
    
    # ===== 为每个议题提取特定会前准备 =====
    for i, topic in enumerate(topics):
        # 首先尝试提取与当前议题直接关联的准备事项
        topic_preparations = []
        
        # 针对当前议题的特定准备事项模式（按精确度排序）
        specific_prep_patterns = [
            # 匹配测试文本中的格式："一、议题，由负责人负责，需要准备事项"
            rf'{i+1}[、.]\s*{re.escape(topic["topic"])}，由{re.escape(topic["leader"])}负责，需要准备([^，,。\n\r;；]+)',
            rf'{i+1}[是、.]\s*{re.escape(topic["topic"])}，由{re.escape(topic["leader"])}负责，需要准备([^，,。\n\r;；]+)',
            # 匹配测试文本中的格式："一、议题，由负责人负责，准备事项"
            rf'{i+1}[、.]\s*{re.escape(topic["topic"])}，由{re.escape(topic["leader"])}负责，准备([^，,。\n\r;；]+)',
            rf'{i+1}[是、.]\s*{re.escape(topic["topic"])}，由{re.escape(topic["leader"])}负责，准备([^，,。\n\r;；]+)',
            # 严格关联议题编号和负责人的模式
            rf'{i+1}[、.]\s*[^，,。\n\r]*?，([^，,。\n\r]+?)需准备',
            rf'{i+1}[、.]\s*[^，,。\n\r]*?，需准备([^，,。\n\r]+?)',
            rf'{i+1}[是、.]\s*[^，,。\n\r]*?，([^，,。\n\r]+?)需准备',
            rf'{i+1}[是、.]\s*[^，,。\n\r]*?，需准备([^，,。\n\r]+?)',
            rf'{topic["leader"]}[你您]?需准备([^，,。\n\r]+?)',
            rf'{topic["leader"]}[你您]?准备([^，,。\n\r]+?)',
            # 匹配议题后紧跟着的准备事项，避免匹配到其他议题
            rf'{i+1}[、.]\s*{re.escape(topic["topic"])}.*?，([^，,。\n\r]+?)需准备',
            rf'{i+1}[、.]\s*{re.escape(topic["topic"])}.*?，需准备([^，,。\n\r]+?)',
            rf'{i+1}[、.]\s*{re.escape(topic["topic"])}.*?，([^，,。\n\r]+?)准备',
            rf'{i+1}[是、.]\s*{re.escape(topic["topic"])}.*?，([^，,。\n\r]+?)需准备',
            rf'{i+1}[是、.]\s*{re.escape(topic["topic"])}.*?，需准备([^，,。\n\r]+?)',
            rf'{i+1}[是、.]\s*{re.escape(topic["topic"])}.*?，([^，,。\n\r]+?)准备',
            # 使用更精确的匹配，避免跨议题匹配
            rf'{i+1}[、.]\s*{re.escape(topic["topic"])}[^，,。\n\r]*?由[^，,。\n\r]+负责[^，,。\n\r]*?，([^，,。\n\r]+?)(?:需要|需)?准备',
            rf'{i+1}[是、.]\s*{re.escape(topic["topic"])}[^，,。\n\r]*?由[^，,。\n\r]+负责[^，,。\n\r]*?，([^，,。\n\r]+?)(?:需要|需)?准备',
            rf'{i+1}[、.]\s*{re.escape(topic["topic"])}[^，,。\n\r]*?由[^，,。\n\r]+负责[^，,。\n\r]*?，需要准备([^，,。\n\r]+?)',
            rf'{i+1}[是、.]\s*{re.escape(topic["topic"])}[^，,。\n\r]*?由[^，,。\n\r]+负责[^，,。\n\r]*?，需要准备([^，,。\n\r]+?)',
            rf'{i+1}[、.]\s*{re.escape(topic["topic"])}[^，,。\n\r]*?由[^，,。\n\r]+负责[^，,。\n\r]*?，准备([^，,。\n\r]+?)',
            rf'{i+1}[是、.]\s*{re.escape(topic["topic"])}[^，,。\n\r]*?由[^，,。\n\r]+负责[^，,。\n\r]*?，准备([^，,。\n\r]+?)',
            # 更通用的模式，但限制匹配范围
            rf'{i+1}[、.]\s*{re.escape(topic["topic"])}[^，,。\n\r]*?，([^，,。\n\r]+?)(?:需要|需)?准备',
            rf'{i+1}[是、.]\s*{re.escape(topic["topic"])}[^，,。\n\r]*?，([^，,。\n\r]+?)(?:需要|需)?准备',
            rf'{i+1}[、.]\s*{re.escape(topic["topic"])}[^，,。\n\r]*?，需要准备([^，,。\n\r]+?)',
            rf'{i+1}[是、.]\s*{re.escape(topic["topic"])}[^，,。\n\r]*?，需要准备([^，,。\n\r]+?)',
            rf'{i+1}[、.]\s*{re.escape(topic["topic"])}[^，,。\n\r]*?，准备([^，,。\n\r]+?)',
            rf'{i+1}[是、.]\s*{re.escape(topic["topic"])}[^，,。\n\r]*?，准备([^，,。\n\r]+?)',
            # 新增更精确的模式，直接匹配"需要准备X"格式
            rf'{i+1}[、.]\s*{re.escape(topic["topic"])}.*?由[^，,。\n\r]+负责.*?需要准备([^，,。\n\r]+)',
            rf'{i+1}[是、.]\s*{re.escape(topic["topic"])}.*?由[^，,。\n\r]+负责.*?需要准备([^，,。\n\r]+)',
            # 新增更简单的模式，直接匹配"准备X"格式
            rf'{i+1}[、.]\s*{re.escape(topic["topic"])}.*?由[^，,。\n\r]+负责.*?准备([^，,。\n\r]+)',
            rf'{i+1}[是、.]\s*{re.escape(topic["topic"])}.*?由[^，,。\n\r]+负责.*?准备([^，,。\n\r]+)',
            # 新增更精确的模式，匹配"由X负责，需要准备Y"格式
            rf'{i+1}[、.]\s*{re.escape(topic["topic"])}.*?由[^，,。\n\r]+负责，需要准备([^，,。\n\r]+)',
            rf'{i+1}[是、.]\s*{re.escape(topic["topic"])}.*?由[^，,。\n\r]+负责，需要准备([^，,。\n\r]+)',
            # 新增更精确的模式，匹配"由X负责，准备Y"格式
            rf'{i+1}[、.]\s*{re.escape(topic["topic"])}.*?由[^，,。\n\r]+负责，准备([^，,。\n\r]+)',
            rf'{i+1}[是、.]\s*{re.escape(topic["topic"])}.*?由[^，,。\n\r]+负责，准备([^，,。\n\r]+)',
            # 新增更简单的模式，直接匹配"需要准备X"格式 - 移除这些通用模式，因为它们会匹配到其他议题的准备事项
            # rf'需要准备([^，,。\n\r]+)',
            # 新增更简单的模式，直接匹配"准备X"格式 - 移除这些通用模式，因为它们会匹配到其他议题的准备事项
            # rf'准备([^，,。\n\r]+)',
            # 新增更精确的模式，匹配议题编号和负责人，然后提取准备事项
            rf'{i+1}[、.]\s*{re.escape(topic["topic"])}，由{re.escape(topic["leader"])}负责，需要准备([^，,。\n\r]+)',
            rf'{i+1}[是、.]\s*{re.escape(topic["topic"])}，由{re.escape(topic["leader"])}负责，需要准备([^，,。\n\r]+)',
            rf'{i+1}[、.]\s*{re.escape(topic["topic"])}，由{re.escape(topic["leader"])}负责，准备([^，,。\n\r]+)',
            rf'{i+1}[是、.]\s*{re.escape(topic["topic"])}，由{re.escape(topic["leader"])}负责，准备([^，,。\n\r]+)'
        ]
        
        # 尝试特定准备事项模式
        found_specific = False
        for pattern in specific_prep_patterns:
            match = safe_re_search(pattern, meeting_text)
            if match:
                preparation_text = match.group(1).strip()
                # 移除可能的标点符号
                preparation_text = safe_re_sub(r'[，,；;.。]$', '', preparation_text)
                # 检查准备事项是否包含其他议题的编号，避免跨议题匹配
                if not any(f'第{j+1}' in preparation_text or f'{j+1}、' in preparation_text or f'{j+1}是' in preparation_text 
                          for j in range(len(topics)) if j != i):
                    topic_preparations.append(preparation_text)
                    print(f"为议题{i+1}提取到特定准备事项: {preparation_text}")
                    found_specific = True
                    break  # 找到第一个合适的匹配就停止
        
        # 如果没有找到特定准备事项，基于会议类型添加通用准备事项
        if not found_specific:
            # 根据会议类型的通用准备模式
            type_specific_patterns = []
            if meeting_info['meeting_type'] == '技术会议':
                type_specific_patterns = [rf'技术.*?{topic["topic"]}.*?准备([^，,。\n\r]+)']
            elif meeting_info['meeting_type'] == '商务会议':
                type_specific_patterns = [rf'商务.*?{topic["topic"]}.*?准备([^，,。\n\r]+)']
            elif meeting_info['meeting_type'] == '项目会议':
                type_specific_patterns = [rf'项目.*?{topic["topic"]}.*?准备([^，,。\n\r]+)']
            elif meeting_info['meeting_type'] == '团队会议':
                type_specific_patterns = [rf'团队.*?{topic["topic"]}.*?准备([^，,。\n\r]+)']
            
            for pattern in type_specific_patterns:
                match = safe_re_search(pattern, meeting_text)
                if match:
                    preparation_text = match.group(1).strip()
                    preparation_text = safe_re_sub(r'[，,；;.。]$', '', preparation_text)
                    topic_preparations.append(preparation_text)
                    print(f"为议题{i+1}提取到类型相关准备事项: {preparation_text}")
                    break
        
        # 设置议题的准备事项
        if topic_preparations:
            topic["preparation"] = '；'.join(topic_preparations)
        elif global_preparations and i == 0:  # 仅将全局准备事项添加到第一个议题
            topic["preparation"] = '；'.join(global_preparations)
            print(f"为第一个议题添加全局准备事项: {topic['preparation']}")
        # 否则保留默认值 '无'
    
    # 如果议题列表为空，尝试创建默认议题
    if not topics:
        topics.append({
            'topic': '会议主要议题',
            'leader': '未指定',
            'preparation': '无'
        })
    
    # 如果没有明确的主题，使用第一个议题作为主题
    if meeting_info['theme'] == '未指定' and topics:
        meeting_info['theme'] = topics[0]['topic'][:50]  # 限制长度
        print(f"推断主题: {meeting_info['theme']}")
    
    # 根据会议类型调整默认主题
    if meeting_info['theme'] == '未指定':
        if meeting_info['meeting_type'] == '技术会议':
            meeting_info['theme'] = '技术讨论会议'
        elif meeting_info['meeting_type'] == '商务会议':
            meeting_info['theme'] = '商务洽谈会议'
        elif meeting_info['meeting_type'] == '项目会议':
            meeting_info['theme'] = '项目进度会议'
        elif meeting_info['meeting_type'] == '团队会议':
            meeting_info['theme'] = '团队例会'
    
    meeting_info['topics'] = topics[:10]  # 限制最多10个议题
    print(f"最终会议信息: {meeting_info}")
    return meeting_info

@app.route("/extract", methods=["POST"])
def extract_meeting_info():
    """提取会议信息的API端点"""
    data = request.json
    text = data.get("text", "")
    if not text:
        return jsonify({"error": "请提供会议文本内容"}), 400
    
    # 首先尝试直接解析输入文本
    print("=== 尝试直接解析输入文本 ===")
    direct_parsed_info = parse_meeting_info(text)
    
    # 检查直接解析的结果质量
    filled_fields = sum(1 for field in ['theme', 'host', 'location', 'attendees', 'duration'] 
                       if direct_parsed_info[field] != '未指定')
    
    # 如果直接解析结果较好，直接返回
    if filled_fields >= 3 or len(direct_parsed_info['topics']) > 0:
        print("直接解析结果质量较好，返回直接解析结果")
        return jsonify({
            "success": True,
            "data": direct_parsed_info,
            "raw_text": text
        })
    
    # 否则使用Ollama提取会议信息
    print("直接解析结果不理想，尝试使用Ollama提取")
    prompt = f"""
你是一名专业的会议信息提取助手，请从以下会议内容中提取关键信息。请仔细分析会议文本，识别各种自然语言表述方式，并严格按照指定格式输出。

## 提取格式要求
请严格按照以下格式输出提取结果，不要添加任何额外说明：
会议主题：具体主题内容
主持人：具体主持人姓名
会议地点：具体地点
参会人员：具体人员名单
会议时长：具体时长
议题1：第一个议题内容
负责人：第一个议题负责人
会前准备：第一个议题的会前准备
议题2：第二个议题内容（如果有）
负责人：第二个议题负责人（如果有）
会前准备：第二个议题的会前准备（如果有）
议题3：第三个议题内容（如果有）
负责人：第三个议题负责人（如果有）
会前准备：第三个议题的会前准备（如果有）

## 注意事项
1. 请识别各种自然语言表述方式，如：
   - 会议主题可能以"关于...的会议"、"讨论..."、"议题是..."等方式表达
   - 地点可能以"在...开个会"、"会议将在...举行"、"地点是..."等方式表达
   - 参会人员可能以"参会的有..."、"参加人员包括..."、"到场的有..."等方式表达
   - 议题可能以"一、二、三"或"1、2、3"编号，或直接列出
   - 负责人可能以"由...负责"、"...来负责"、"...牵头"等方式表达

2. 如果某个字段不存在，请留空但保留字段名。
3. 请提取完整的人员姓名，不要缩写。
4. 请尽可能准确地识别所有议题，即使它们没有明确编号。
5. 对于会前准备，请提取具体需要准备的事项，如资料、文件、演示等。

## 示例理解
例如，对于文本："明天上午10点在三楼会议室开个关于项目进度的会议，参加的有张三、李四和王五，会议大概开1小时，主要讨论三个议题：一是项目当前进度，由张三负责；二是遇到的问题，由李四准备相关资料；三是下一步计划，王五负责。"，应该提取为：
会议主题：项目进度
主持人：
会议地点：三楼会议室
参会人员：张三、李四、王五
会议时长：1小时
议题1：项目当前进度
负责人：张三
会前准备：
议题2：遇到的问题
负责人：李四
会前准备：准备相关资料
议题3：下一步计划
负责人：王五
会前准备：

会议内容：
{text}
"""
    
    try:
        # 调用本地模型
        result = call_ollama(prompt)
        print(f"Ollama响应: {result}")
        
        # 解析Ollama提取的信息
        ollama_parsed_info = parse_meeting_info(result)
        
        # 合并结果，优先使用Ollama提取的信息
        final_info = direct_parsed_info.copy()
        for field in ['theme', 'host', 'location', 'attendees', 'duration']:
            if ollama_parsed_info[field] != '未指定':
                final_info[field] = ollama_parsed_info[field]
        
        # 如果Ollama提取了议题，则使用Ollama的议题
        if len(ollama_parsed_info['topics']) > 0:
            final_info['topics'] = ollama_parsed_info['topics']
        
        return jsonify({
            "success": True,
            "data": final_info,
            "raw_text": result
        })
    except Exception as e:
        print(f"使用Ollama提取失败: {str(e)}")
        # 即使Ollama失败，也返回直接解析的结果
        return jsonify({
            "success": True,
            "data": direct_parsed_info,
            "raw_text": text
        })


@app.route("/generate", methods=["POST"])
def generate_meeting_doc():
    """生成会议纪要 Word 文件"""
    data = request.json
    text = data.get("text", "")
    if not text:
        return jsonify({"error": "请提供会议文本内容"}), 400

    # 直接使用 parse_meeting_info 解析，不再调用 Ollama
    print("=== 开始生成文档，直接解析输入文本 ===")
    meeting_info = parse_meeting_info(text)
    
    # 如果直接解析结果不理想，再尝试使用 Ollama
    filled_fields = sum(1 for field in ['theme', 'host', 'location', 'attendees', 'duration'] 
                       if meeting_info[field] != '未指定')
    
    if filled_fields < 3 and len(meeting_info['topics']) == 0:
        print("直接解析结果不理想，尝试使用 Ollama")
        extract_prompt = f"""
你是一名专业的会议信息提取助手，请从以下会议内容中提取关键信息。请仔细分析会议文本，识别各种自然语言表述方式，并严格按照指定格式输出。

## 提取格式要求
请严格按照以下格式输出提取结果，不要添加任何额外说明：
会议主题：具体主题内容
主持人：具体主持人姓名
会议地点：具体地点
参会人员：具体人员名单
会议时长：具体时长
议题1：第一个议题内容
负责人：第一个议题负责人
会前准备：第一个议题的会前准备
议题2：第二个议题内容（如果有）
负责人：第二个议题负责人（如果有）
会前准备：第二个议题的会前准备（如果有）

会议内容：
{text}
"""
        extracted_text = call_ollama(extract_prompt)
        meeting_info = parse_meeting_info(extracted_text)
    
    # 确保输出文件夹存在
    output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
    os.makedirs(output_dir, exist_ok=True)

    # 文件路径
    filename = f"meeting_{int(time.time())}.docx"
    file_path = os.path.join(output_dir, filename)

    # 检查是否有会前准备事项
    try:
        # 安全地提取会前准备事项
        def safe_re_search(pattern, text, flags=0):
            try:
                return re.search(pattern, text, flags)
            except re.error as e:
                print(f"正则表达式错误: {pattern}, {e}")
                return None
        
        preparation_items_match = safe_re_search(r'会前准备事项[：:]\s*(.+)', extracted_text, re.DOTALL)
        if preparation_items_match:
            meeting_info['preparation_items'] = preparation_items_match.group(1).strip()
    except Exception as e:
        print(f"提取会前准备事项异常: {e}")
    
    try:
        # 使用专业的Word生成器创建文档
        generator = WordMeetingGenerator()
        file_path = generator.create_meeting_document(meeting_info, file_path)

        print(f"✅ 会议纪要已保存: {file_path}")
        
        # 验证文件是否存在
        if not os.path.exists(file_path):
            print(f"❌ 文件不存在: {file_path}")
            return jsonify({"error": "文档生成失败，文件未创建"}), 500

        # 返回文件，添加Content-Disposition头
        return send_file(
            file_path, 
            as_attachment=True, 
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            download_name=f"meeting_{int(time.time())}.docx"
        )
    except Exception as e:
        print(f"❌ 文档生成异常: {str(e)}")
        return jsonify({"error": f"文档生成失败: {str(e)}"}), 500

@app.route("/health", methods=["GET"])
def health_check():
    """健康检查端点"""
    return jsonify({"status": "healthy", "message": "服务运行正常"})

if __name__ == "__main__":
    app.run(port=5000, debug=True)
