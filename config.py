from pathlib import Path
import loguru

base_dir = Path(__file__).resolve().parent

# 设置登录信息文件路径
login_info_path = base_dir / Path(r'db\login_info.txt')

# 设置日志文件路径
app_log_path = base_dir / Path(r'log/app.log')
error_log_path = base_dir / Path(r'log/errors.log')
# app_log_path = os.path.join(os.path.dirname(__file__), "log/app.log")
# error_log_path = os.path.join(os.path.dirname(__file__), "log/errors.log")

# 配置应用日志格式和输出，逐日备份
app_logger = loguru.logger
app_logger.add(app_log_path, format="{time} | {level} | {message}",
               filter=lambda record: record["level"].name == "INFO" or record["level"].name == "WARNING",
               rotation="1 day", retention="60 days")

# 配置错误日志格式和输出，逐日备份
error_logger = loguru.logger
error_logger.add(error_log_path, format="{time} | {level} | {message}",
                 filter=lambda record: record["level"].name == "ERROR" or record["level"].name == "CRITICAL",
                 rotation="1 day", retention="60 days")

source_file_path = r'D:\网盘同步\1.药事\3.抗菌药物监测\2025年'
freq_web_list = ['即刻', '1/日', '2/日', '3/日', '4/日', 'q2h', 'q6h', 'q8h', 'q12h', '每晚', '其他']
way_web_list = ['静脉滴注', '静脉泵入', '静脉推注', '肌肉注射', '静脉注射', '皮下注射', '球后注射',
                '结膜下注射', '眼内注射', '直肠给药', '雾化吸入', '肠道准备',
                '口服', '外用', '滴鼻', '滴耳', '滴眼', '鞘内注射', '腹膜透析', '皮试']
dose_unit_web_list = ['克', '毫克', '万单位', '滴', 'ml', '片', '支', '粒', '瓶', '包', '袋']

freq_dict = {'ONCE': '即刻',
             'QD': '1/日',
             'BID': '2/日',
             'TID': '3/日',
             'QID': '4/日',
             'Q2H': 'q2h',
             'Q6H': 'q6h',
             'Q8H': 'q8h',
             'Q12H': 'q12h',
             'QN': '每晚',
             '(空白)': '其他'}
way_dict = {'口服': '口服',
            '血液透析 ': '腹膜透析',
            '外用': '外用',
            '肌肉注射': '肌肉注射',
            '吸入': '雾化吸入',
            '涂眼睑内': '外用',
            '滴眼': '滴眼',
            '皮下注射(不带费用、耗材）': '皮下注射',
            '静脉注射(麻醉科专用)': '静脉注射',
            '塞肛': '外用',
            '含服': '口服',
            '局麻用': '肌肉注射',
            '喷喉': '外用'}
dose_unit_dict = {'丸': '粒', 'ug': '粒', '吸': '粒', 'ml': 'ml', '片': '片', 'g': '克', '粒': '粒', 'mg': '毫克'}
