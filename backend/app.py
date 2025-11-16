from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import subprocess
import re
import os
import time
from word_generator import WordMeetingGenerator

app = Flask(__name__)
CORS(app)

class RegexHelper:
    """正则表达式辅助类"""
    @staticmethod
    def safe_search(pattern, text, flags=0):
        try:
            return re.search(pattern, text, flags)
        except re.error as e:
            print(f"正则表达式错误: {pattern}, {e}")
            return None
    
    @staticmethod
    def safe_sub(pattern, repl, text, flags=0):
        try:
            return re.sub(pattern, repl, text, flags)
        except re.error as e:
            print(f"正则表达式替换错误: {e}")
            return text
    
    @staticmethod
    def safe_split(pattern, text, maxsplit=0, flags=0):
        try:
            return re.split(pattern, text, maxsplit=maxsplit, flags=flags)
        except re.error as e:
            print(f"正则表达式分割错误: {e}")
            return [text]
    
    @staticmethod
    def safe_findall(pattern, text, flags=0):
        try:
            return re.findall(pattern, text, flags=flags)
        except re.error as e:
            print(f"正则表达式查找错误: {e}")
            return []

def call_ollama(prompt, model="llama3"):
    """调用本地 Ollama 模型"""
    try:
        result = subprocess.run(
            ["ollama", "run", model],
            input=prompt.encode("utf-8"),
            capture_output=True,
            timeout=120
        )
        output = result.stdout.decode("utf-8", errors="ignore")
        if output:
            output = RegexHelper.safe_sub(r'^.*?---', '', output, re.DOTALL)
        return output.strip() if output else ""
    except subprocess.TimeoutExpired:
        return "模型响应超时，请检查 Ollama 是否正常运行。"
    except Exception as e:
        print(f"调用Ollama详细错误: {e}")
        return f"调用 Ollama 出错：{str(e)}"

def detect_meeting_type(text):
    """检测会议类型"""
    keywords = {
        '技术会议': ['技术', '开发', '编程', '代码', 'API', '架构', '数据库', '测试', 'bug', '部署', '性能', '优化'],
        '商务会议': ['商务', '合作', '谈判', '客户', '市场', '销售', '营销', '推广', '策略', '预算', '财务'],
        '项目会议': ['项目', '进度', '里程碑', '任务', '分工', '责任', '延期', '风险', '协调', '资源'],
        '团队会议': ['团队', '部门', '周会', '例会', '分享', '讨论', '交流', '培训', '总结', '回顾']
    }
    
    keyword_count = {
        meeting_type: sum(keyword in text for keyword in kw_list)
        for meeting_type, kw_list in keywords.items()
    }
    
    max_count = max(keyword_count.values())
    if max_count > 1:
        return max(keyword_count.items(), key=lambda x: x[1])[0]
    return '通用会议'

def extract_field(text, patterns):
    """通用字段提取函数"""
    for pattern in patterns:
        match = RegexHelper.safe_search(pattern, text)
        if match:
            return match.group(1).strip()
    return None

def extract_attendees(text):
    """提取参会人员"""
    patterns = [
        r'参会(?:人员)?[：:]\s*([^。\n\r]+?)(?:。|会议)',
        r'参会的有([^。\n\r]+?)(?:。|会议)',
        r'参加(?:人员)?包括([^。\n\r]+?)(?:。|会议)',
        r'出席(?:人员)?[：:]\s*([^。\n\r]+?)(?:。|会议)'
    ]
    
    attendees_text = extract_field(text, patterns)
    if not attendees_text:
        return '未指定'
    
    # 清理和分割人员名单
    attendees = []
    for person in RegexHelper.safe_split(r'[，,、]', attendees_text):
        person = person.strip()
        # 过滤无关文本
        if person and not any(x in person for x in ['等', '会议主要', '大概', '开', '小时']):
            # 保留部门信息
            person = RegexHelper.safe_sub(r'还有新来的', '', person)
            person = RegexHelper.safe_sub(r'还有', '', person)
            person = RegexHelper.safe_sub(r'以及', '', person)
            person = person.strip()
            if person:
                attendees.append(person)
    
    return '，'.join(list(dict.fromkeys(attendees))) if attendees else '未指定'

def extract_topics(text, meeting_type):
    """提取会议议题"""
    topics = []
    
    # 尝试匹配中文数字格式
    chinese_nums = ['一', '二', '三', '四', '五', '六', '七', '八', '九', '十']
    for i, num in enumerate(chinese_nums, 1):
        patterns = [
            rf'{num}[是、.]\s*(.+?)[，,；；。。\n\r]',
            rf'{num}[是、.]\s*(.+?)$'
        ]
        
        topic_text = extract_field(text, patterns)
        if topic_text:
            topics.append({
                'topic': topic_text,
                'leader': '未指定',
                'preparation': '无'
            })
    
    return topics

def extract_leader_for_topic(text, topic, index):
    """为议题提取负责人"""
    chinese_nums = ['一', '二', '三', '四', '五', '六', '七', '八', '九', '十']
    num = chinese_nums[index] if index < len(chinese_nums) else str(index+1)
    
    # 提取议题相关的文本片段（从当前议题到下一个议题或句号）
    topic_pattern = rf'{num}是[^。；;]+?(?=[二三四五六七八九十]是|。|；|$)'
    topic_match = RegexHelper.safe_search(topic_pattern, text)
    topic_segment = topic_match.group(0) if topic_match else text
    
    patterns = [
        # 匹配 "XX你准备" 或 "XX负责"
        r'([^，,。；;\n\r]{2,4})(?:你|您)(?:准备|负责|牵头)',
        # 匹配 "XX负责这块"
        r'([^，,。；;\n\r]{2,4})负责(?:这块|该议题)',
        # 匹配 "由XX负责"
        r'由([^，,。；;\n\r]{2,4})负责'
    ]
    
    leader = extract_field(topic_segment, patterns)
    if leader:
        # 清理负责人名称
        leader = RegexHelper.safe_sub(r'[你您]|负责|牵头|主讲|汇报|^由|特别是.*|，.*', '', leader).strip()
        return leader if leader else '未指定'
    return '未指定'

def extract_preparation_for_topic(text, topic, index):
    """为议题提取会前准备"""
    leader = topic.get("leader", "")
    if leader and leader != '未指定':
        # 提取议题相关的文本片段
        chinese_nums = ['一', '二', '三', '四', '五', '六', '七', '八', '九', '十']
        num = chinese_nums[index] if index < len(chinese_nums) else str(index+1)
        topic_pattern = rf'{num}是[^。；;]+?(?=[二三四五六七八九十]是|。|；|$)'
        topic_match = RegexHelper.safe_search(topic_pattern, text)
        topic_segment = topic_match.group(0) if topic_match else text
        
        patterns = [
            rf'{re.escape(leader)}[你您]?准备下?([^，,。；;\n\r]+?)(?:[，,。；;]|$)',
            rf'{re.escape(leader)}[你您]?需准备([^，,。；;\n\r]+?)(?:[，,。；;]|$)'
        ]
        
        preparation = extract_field(topic_segment, patterns)
        if preparation:
            # 清理准备事项
            preparation = RegexHelper.safe_sub(r'；.*|二是.*|三是.*', '', preparation)
            preparation = RegexHelper.safe_sub(r'[，,；;.。]$', '', preparation)
            return preparation.strip() if preparation.strip() else '无'
    return '无'

def parse_meeting_info(meeting_text):
    """解析会议信息"""
    if not meeting_text or not isinstance(meeting_text, str):
        print("无效的会议文本")
        return {
            'theme': '未指定',
            'host': '未指定',
            'location': '未指定',
            'attendees': '未指定',
            'duration': '未指定',
            'topics': [],
            'meeting_type': '通用会议'
        }
    
    print("解析文本:", repr(meeting_text[:100]))
    
    # 检测会议类型
    meeting_type = detect_meeting_type(meeting_text)
    print(f"检测到会议类型: {meeting_type}")
    
    # 提取基本信息
    host_patterns = [
        r'主持人[：:]\s*(.+?)[\n\r]',
        r'主持人[：:]\s*(.+)',
        r'由([^，,。\n\r]+)主持'
    ]
    
    location_patterns = [
        r'(?:在|于)\s*([^，,。\n\r]+?)\s*(?:召开|举行|开)',
        r'会议地点[：:]\s*(.+?)[\n\r]',
        r'会议地点[：:]\s*(.+)'
    ]
    
    duration_patterns = [
        r'会议大概开([^，,。\n\r]+)',
        r'会议时长[：:]\s*(.+?)[\n\r]',
        r'会议时长[：:]\s*(.+)',
        r'大约([^，,。\n\r]+)小时'
    ]
    
    theme_patterns = [
        r'会议主题[：:]\s*(.+?)[\n\r]',
        r'会议主题[：:]\s*(.+)',
        r'讨论([^，,。\n\r]+)',
        r'关于([^，,。\n\r]+)的会议'
    ]
    
    meeting_info = {
        'theme': extract_field(meeting_text, theme_patterns) or '未指定',
        'host': extract_field(meeting_text, host_patterns) or '未指定',
        'location': extract_field(meeting_text, location_patterns) or '未指定',
        'attendees': extract_attendees(meeting_text),
        'duration': extract_field(meeting_text, duration_patterns) or '未指定',
        'meeting_type': meeting_type,
        'topics': []
    }
    
    # 提取议题
    topics = extract_topics(meeting_text, meeting_type)
    
    # 为每个议题提取负责人和准备事项
    for i, topic in enumerate(topics):
        topic['leader'] = extract_leader_for_topic(meeting_text, topic, i)
        topic['preparation'] = extract_preparation_for_topic(meeting_text, topic, i)
        print(f"议题{i+1}: {topic['topic']}, 负责人: {topic['leader']}")
    
    meeting_info['topics'] = topics[:10]  # 限制最多10个议题
    
    # 如果没有明确主题，使用第一个议题
    if meeting_info['theme'] == '未指定' and topics:
        meeting_info['theme'] = topics[0]['topic'][:50]
    
    print(f"最终会议信息: {meeting_info}")
    return meeting_info

@app.route("/extract", methods=["POST"])
def extract_meeting_info():
    """提取会议信息的API端点"""
    data = request.json
    text = data.get("text", "")
    if not text:
        return jsonify({"error": "请提供会议文本内容"}), 400
    
    print("=== 尝试直接解析输入文本 ===")
    direct_parsed_info = parse_meeting_info(text)
    
    # 检查解析质量
    filled_fields = sum(1 for field in ['theme', 'host', 'location', 'attendees', 'duration'] 
                       if direct_parsed_info[field] != '未指定')
    
    if filled_fields >= 3 or len(direct_parsed_info['topics']) > 0:
        print("直接解析结果质量较好")
        return jsonify({
            "success": True,
            "data": direct_parsed_info,
            "raw_text": text
        })
    
    # 使用Ollama作为后备
    print("直接解析结果不理想，尝试使用Ollama")
    try:
        prompt = f"""请从以下会议内容中提取关键信息，按以下格式输出：
会议主题：
主持人：
会议地点：
参会人员：
会议时长：
议题1：
负责人：
会前准备：

会议内容：
{text}
"""
        result = call_ollama(prompt)
        ollama_parsed_info = parse_meeting_info(result)
        
        # 合并结果
        final_info = direct_parsed_info.copy()
        for field in ['theme', 'host', 'location', 'attendees', 'duration']:
            if ollama_parsed_info[field] != '未指定':
                final_info[field] = ollama_parsed_info[field]
        
        if len(ollama_parsed_info['topics']) > 0:
            final_info['topics'] = ollama_parsed_info['topics']
        
        return jsonify({
            "success": True,
            "data": final_info,
            "raw_text": result
        })
    except Exception as e:
        print(f"使用Ollama提取失败: {str(e)}")
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

    print("=== 开始生成文档 ===")
    meeting_info = parse_meeting_info(text)
    
    # 确保输出文件夹存在
    output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
    os.makedirs(output_dir, exist_ok=True)

    filename = f"meeting_{int(time.time())}.docx"
    file_path = os.path.join(output_dir, filename)
    
    try:
        generator = WordMeetingGenerator()
        file_path = generator.create_meeting_document(meeting_info, file_path)
        print(f"✅ 会议纪要已保存: {file_path}")
        
        if not os.path.exists(file_path):
            return jsonify({"error": "文档生成失败"}), 500

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
