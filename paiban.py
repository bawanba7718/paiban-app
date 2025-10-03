import pandas as pd
import datetime
from datetime import datetime, time, timedelta
import streamlit as st
import openpyxl
from openpyxl import load_workbook
import os
import time as t
import tempfile
from webdav3.client import Client
import threading
import html

class AgentViewer:
    def __init__(self):
        # 颜色-职位对应关系 - 扩展颜色映射
        self.color_roles = {
            'FFC000': 'C席',
            'FFEE79': 'C席',
            'E2EFDA': 'C席',
            '91AADF': 'C席',
            'D9E1F2': 'C席',  # 修正为D9E1F2
            'EF949F': 'B席',
            'FADADE': 'B席',
            '8CDDFA': '休',
            'FFFF00': '休',
            'FFFFFF': 'A席',  # 默认白色为A席
        }
        
        # 席位颜色映射
        self.seat_colors = {
            'C席': '#FFC000',
            'B席': '#EF949F',
            'A席': '#FFFFFF'
        }
        
        # 状态图标
        self.status_icons = {
            '搬砖中': '🛠️',
            '干饭中': '🍚',
            '已回家': '🏠',
            '正在路上': '🚗',
            '未排班': '❓',
            '未知班次': '❓'
        }
        
        # 班次时间定义 - 修正B席班次时间
        self.shift_times = {
            'T1': {'start': time(8, 0), 'end': time(20, 0), 'name': '白班', 
                  'break_start': time(13, 0), 'break_end': time(14, 0)},
            'T2': {'start': time(20, 0), 'end': time(8, 0), 'name': '夜班',
                  'break_start': None, 'break_end': None},
            'M2': {'start': time(8, 0), 'end': time(17, 0), 'name': '早班',
                  'break_start': time(14, 0), 'break_end': time(15, 0)},
            'E2': {'start': time(13, 0), 'end': time(22, 0), 'name': '晚班',
                  'break_start': time(15, 0), 'break_end': time(16, 0)},
            'E3': {'start': time(13, 0), 'end': time(23, 0), 'name': '晚班',
                  'break_start': time(15, 0), 'break_end': time(17, 0)},
            'M1': {'start': time(7, 0), 'end': time(16, 0), 'name': '早班',
                  'break_start': time(12, 0), 'break_end': time(13, 0)},
            'D1': {'start': time(9, 0), 'end': time(18, 0), 'name': '白班',
                  'break_start': time(12, 0), 'break_end': time(13, 0)},
            'D2': {'start': time(10, 0), 'end': time(19, 0), 'name': '白班',
                  'break_start': time(13, 0), 'break_end': time(14, 0)},
            'D3': {'start': time(11, 0), 'end': time(20, 0), 'name': '白班',
                  'break_start': time(15, 0), 'break_end': time(16, 0)},
            'E1': {'start': time(12, 0), 'end': time(21, 0), 'name': '晚班',
                  'break_start': time(15, 0), 'break_end': time(16, 0)},
            'F1': {'start': time(7, 0), 'end': time(13, 0), 'name': '短班',
                  'break_start': None, 'break_end': None},
            'F2': {'start': time(9, 0), 'end': time(16, 0), 'name': '短班',
                  'break_start': time(13, 0), 'break_end': time(14, 0)},
            'F3': {'start': time(17, 0), 'end': time(23, 0), 'name': '短班',
                  'break_start': None, 'break_end': None},
            'H1': {'start': time(7, 0), 'end': time(11, 0), 'name': '半日班',
                  'break_start': None, 'break_end': None},
            'H2': {'start': time(16, 0), 'end': time(20, 0), 'name': '半日班',
                  'break_start': None, 'break_end': None},
        }

    def get_work_status(self, shift_code, seat, check_time=None):
        if not shift_code or str(shift_code).strip() == '':
            return "未排班", "#BFBFBF"
            
        shift_code = str(shift_code).strip()
        main_shift = None
        
        # 改进班次识别逻辑，处理包含"备"的班次
        for s in self.shift_times:
            if s in shift_code:
                main_shift = s
                break
                
        if not main_shift:
            return "未知班次", "#BFBFBF"
            
        shift = self.shift_times[main_shift].copy()
        
        # 修正A席T1班次的休息时间
        if seat == 'A席' and main_shift == 'T1':
            shift['break_start'] = time(14, 0)
            shift['break_end'] = time(15, 0)
            
        check_time = check_time or datetime.now().time()
        
        start, end = shift['start'], shift['end']
        break_start, break_end = shift.get('break_start'), shift.get('break_end')
        
        is_night_shift = main_shift == 'T2'
        in_work_time = False
        
        # 改进跨天班次判断逻辑
        if is_night_shift:
            # T2班次：前一天20:00到当天8:00
            if start <= end:  # 正常时间范围
                in_work_time = start <= check_time < end
            else:  # 跨天情况
                in_work_time = (check_time >= start) or (check_time < end)
        else:
            # 非跨天班次
            in_work_time = start <= check_time < end
            
        is_on_the_way = False
        if not in_work_time:
            if not is_night_shift:
                # 非夜班：上班前显示"正在路上"
                is_on_the_way = check_time < start
            else:
                # 夜班：在当天20:00前显示"正在路上"
                if check_time < start and check_time >= time(0, 0):
                    is_on_the_way = True
        
        in_break_time = False
        if break_start and break_end and in_work_time:
            if break_start < break_end:
                in_break_time = break_start <= check_time < break_end
            else:
                # 处理跨天休息时间（目前没有这种情况）
                in_break_time = check_time >= break_start or check_time < break_end
        
        if is_on_the_way:
            return "正在路上", "#BFBFBF"
        elif not in_work_time:
            return "已回家", "#BFBFBF"
        elif in_break_time:
            return "干饭中", "orange"
        else:
            return "搬砖中", "green"

    def get_cell_color(self, cell):
        try:
            if cell and cell.fill and cell.fill.start_color:
                color = cell.fill.start_color.rgb
                if color:
                    color_str = str(color).upper()
                    # 处理颜色格式，确保返回6位颜色码
                    if color_str.startswith('FF'):
                        color_str = color_str[2:]  # 去除alpha通道
                    elif len(color_str) == 8:
                        color_str = color_str[2:]  # 处理8位颜色码
                    return color_str if len(color_str) == 6 else "FFFFFF"
            return "FFFFFF"
        except:
            return "FFFFFF"
    
    def load_schedule_with_colors(self, file_path, target_date):
        try:
            if not os.path.exists(file_path):
                st.error(f"文件不存在: {file_path}")
                return None
                
            wb = load_workbook(file_path, data_only=True)
            if '全部排班' not in wb.sheetnames:
                st.error("工作表 '全部排班' 不存在")
                return None
            
            main_sheet = wb['全部排班']
            df_main = pd.read_excel(file_path, sheet_name='全部排班')
            
            target_date_str = target_date.strftime('%Y-%m-%d')
            today_col_idx = None
            
            # 改进日期列查找逻辑
            for idx, col in enumerate(df_main.columns):
                col_str = str(col)
                if (target_date_str in col_str or 
                    target_date.strftime('%m-%d') in col_str or
                    target_date.strftime('%Y/%m/%d') in col_str or
                    target_date.strftime('%m/%d') in col_str):
                    today_col_idx = idx
                    break
            
            if today_col_idx is None:
                st.error(f"未找到 {target_date_str} 的排班列")
                return None
            
            color_data = []
            for row_idx, row in enumerate(main_sheet.iter_rows(min_row=2, values_only=False), start=2):
                try:
                    if len(row) < 4:
                        continue
                        
                    name = str(row[3].value).strip() if row[3].value else ''
                    if not name:
                        continue
                    
                    shift_code = ""
                    color_code = "FFFFFF"
                    if len(row) > today_col_idx:
                        shift_cell = row[today_col_idx]
                        shift_code = str(shift_cell.value).strip() if shift_cell.value else ""
                        color_code = self.get_cell_color(shift_cell)
                    
                    # 放宽过滤条件，包含"备"的班次也显示
                    if (not shift_code or 
                        shift_code.strip() in ['', '休', '休息']):
                        continue
                    
                    # 改进颜色识别，包含更多颜色变体
                    seat = self.color_roles.get(color_code, 'A席')
                    
                    person_info = {
                        'name': name,
                        'id': str(row[2].value).strip() if row[2].value else '',
                        'workplace': str(row[0].value).strip() if row[0].value else '',
                        'shift': shift_code,
                        'color': color_code,
                        'seat': seat,
                        'status': '',
                        'status_color': ''
                    }
                    
                    color_data.append(person_info)
                except Exception as e:
                    continue
            
            return pd.DataFrame(color_data)
            
        except Exception as e:
            st.error(f"加载数据失败: {str(e)}")
            return None
    
    def categorize_by_seat(self, df, check_time=None):
        result = {'A席': [], 'B席': [], 'C席': []}
        if df is None or df.empty:
            return result
        
        for _, person in df.iterrows():
            status, status_color = self.get_work_status(person['shift'], person['seat'], check_time)
            person['status'] = status
            person['status_color'] = status_color
            
            seat = person['seat']
            if seat in result:
                result[seat].append(person)
            else:
                result['A席'].append(person)
        
        for cat in result:
            result[cat].sort(key=lambda x: 
                0 if x['status'] == '搬砖中' else 
                1 if x['status'] == '干饭中' else 
                2 if x['status'] == '正在路上' else 3)
        
        return result

def download_from_jiananguo():
    try:
        # 从Streamlit Secrets获取坚果云凭证（部署时需配置）
        jiananguo_email = st.secrets.get("JIANANGUO_EMAIL", "")
        jiananguo_password = st.secrets.get("JIANANGUO_PASSWORD", "")
        
        if not jiananguo_email or not jiananguo_password:
            return False, None, "未配置坚果云邮箱或密码（需在Streamlit Secrets中设置）"
        
        options = {
            'webdav_hostname': 'https://dav.jianguoyun.com/dav/',
            'webdav_login': jiananguo_email,
            'webdav_password': jiananguo_password
        }
        
        client = Client(options)
        remote_file = '我的坚果云/排班.xlsx'  # 坚果云中排班文件的路径（需与实际一致）
        
        # 创建临时文件存储下载的排班表
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            local_file = tmp_file.name
        
        # 同步下载文件
        client.download_sync(remote_path=remote_file, local_path=local_file)
        
        if os.path.exists(local_file) and os.path.getsize(local_file) > 0:
            return True, local_file, "成功从坚果云下载排班文件"
        else:
            return False, None, "从坚果云下载文件为空或不存在"
            
    def download_from_jiananguo():
    try:
        # （原有下载逻辑，保持不变）
    except Exception as e:
        import traceback
        error_detail = traceback.format_exc()  # 新增：定义error_detail变量
        return False, None, f"下载失败: {str(e)}\n详细错误:\n{error_detail}"

def create_agent_card(person_info, viewer):
    """创建坐席信息卡片 - 根据状态设置背景色"""
    # 获取状态图标
    status_icon = viewer.status_icons.get(person_info['status'], '❓')
    
    # 统一状态颜色：正在路上和已回家都使用 #BFBFBF
    if person_info['status'] in ["正在路上", "已回家"]:
        status_color = "#BFBFBF"
        bg_color = "#BFBFBF"  # 状态为非工作中时，卡片背景设为灰色
    else:
        status_color = person_info['status_color']
        # 正常工作状态：按席位颜色设置背景
        seat_type = person_info['seat']
        bg_color = f"#{person_info['color']}" if seat_type in ['B席', 'C席'] else "#FFFFFF"
    
    # HTML卡片（支持Streamlit渲染）
    card_html = f"""
    <div style="background-color: {bg_color}; border: 2px solid #000000; border-radius: 8px; padding: 12px; margin: 8px 0; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
        <div style="display: flex; align-items: center; justify-content: space-between;">
            <div style="width: 30%;">
                <p style="margin: 5px 0; font-size: 14px; font-weight: bold;">工号: {person_info['id']}</p>
                <p style="margin: 5px 0; font-size: 14px; font-weight: bold;">职场: {person_info['workplace']}</p>
            </div>
            <div style="width: 40%; text-align: center;">
                <h3 style="margin: 0 0 8px 0; font-size: 18px; font-weight: bold;">{person_info['name']}</h3>
                <span style="font-size: 24px; display: block; margin-bottom: 4px;">{status_icon}</span>
                <p style="margin: 0; font-size: 16px; font-weight: bold; color: {status_color};">{person_info['status']}</p>
            </div>
            <div style="width: 30%; text-align: right;">
                <p style="margin: 5px 0; font-size: 14px; font-weight: bold;">班次: {person_info['shift']}</p>
                <p style="margin: 5px 0; font-size: 14px; font-weight: bold;">席位: {person_info['seat']}</p>
            </div>
        </div>
    </div>
    """
    
    return card_html

def create_stat_card(seat, online_count, total_count, color):
    """创建坐席统计卡片"""
    return f"""
    <div style="background-color: {color}; border: 2px solid #000000; border-radius: 8px; padding: 12px; margin: 8px 0; text-align: center; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
        <h3 style="margin: 0 0 6px 0; font-size: 18px; font-weight: bold;">{seat}</h3>
        <div style="display: flex; justify-content: center; align-items: baseline;">
            <span style="font-size: 26px; font-weight: bold; color: #2E8B57;">{online_count}</span>
            <span style="font-size: 18px; color: black; margin: 0 6px;">/</span>
            <span style="font-size: 22px; font-weight: bold; color: black;">{total_count}</span>
        </div>
        <p style="margin: 4px 0 0 0; color: black; font-size: 14px;">在线/总人数</p>
    </div>
    """

def update_current_time():
    """更新当前时间（含星期）"""
    weekdays = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]
    now = datetime.now()
    weekday = weekdays[now.weekday()]
    return now.strftime(f"%Y年%m月%d日 {weekday} %H:%M:%S")

def auto_refresh_time(placeholder):
    """自动刷新当前时间（每秒更新）"""
    while True:
        if not st.session_state.get('auto_refresh', True):
            t.sleep(1)
            continue
        placeholder.markdown(f"### 当前时间: {update_current_time()}")
        t.sleep(1)

def main():
    # 配置Streamlit页面（标题、布局、图标）
    st.set_page_config(
        page_title="综合组在线坐席", 
        layout="wide",
        page_icon="📊"
    )
    
    # 初始化Session State（存储临时数据，避免刷新丢失）
    if 'file_path' not in st.session_state:
        st.session_state.file_path = None
    if 'last_download' not in st.session_state:
        st.session_state.last_download = None
    if 'last_refresh' not in st.session_state:
        st.session_state.last_refresh = None
    if 'schedule_data' not in st.session_state:
        st.session_state.schedule_data = None
    if 'auto_refresh' not in st.session_state:
        st.session_state.auto_refresh = True
    if 'refresh_counter' not in st.session_state:
        st.session_state.refresh_counter = 0
    if 'last_load_date' not in st.session_state:
        st.session_state.last_load_date = None
    if 'last_auto_refresh' not in st.session_state:
        st.session_state.last_auto_refresh = datetime.now()
    if 'force_current_time' not in st.session_state:
        st.session_state.force_current_time = True  # 强制显示当前时段
    st.session_state.force_current_time = True  # 每次渲染都锁定为True，防止被重置
    
    # 初始化坐席查看器
    viewer = AgentViewer()
    
    # 首次运行或本地文件不存在时，从坚果云下载排班表
    if st.session_state.file_path is None or not os.path.exists(st.session_state.file_path):
        with st.spinner("正在下载排班文件..."):
            download_success, file_path, download_message = download_from_jiananguo()
            if download_success:
                st.session_state.file_path = file_path
                st.session_state.last_download = datetime.now()
                st.success("排班文件下载成功")
            else:
                st.error(f"下载失败: {download_message}")
                st.stop()  # 下载失败则停止程序
    
    # 每小时自动刷新排班数据（避免数据过时）
    current_time = datetime.now()
    time_diff = current_time - st.session_state.last_auto_refresh
    if time_diff.total_seconds() >= 3600:  # 3600秒 = 1小时
        st.session_state.last_auto_refresh = current_time
        st.session_state.refresh_counter += 1
        st.session_state.schedule_data = None  # 清除旧数据缓存
        st.rerun()  # 重新运行程序加载新数据
    
    # 页面标题
    st.title("📊 综合组在线坐席")
    
    # 顶部控制栏（时间显示、刷新、重新下载）
    col1, col2, col3 = st.columns([3, 1, 1])
    
    with col1:
        # 实时更新当前时间（用线程避免阻塞）
        current_datetime = st.empty()
        if 'time_thread' not in st.session_state:
            st.session_state.time_thread = threading.Thread(
                target=auto_refresh_time, 
                args=(current_datetime,), 
                daemon=True  # 线程随主程序退出而关闭
            )
            st.session_state.time_thread.start()
    
    with col2:
        # 手动刷新坐席状态
        if st.button("🔄 刷新状态", use_container_width=True):
            st.session_state.last_refresh = datetime.now()
            st.session_state.refresh_counter += 1
            st.session_state.schedule_data = None  # 清除旧数据
            st.session_state.force_current_time = True  # 刷新后仍强制当前时段
            st.success("状态已刷新")
    
    with col3:
        # 重新下载排班表（手动更新）
        if st.button("📥 重新下载班表", use_container_width=True):
            with st.spinner("重新下载班表中..."):
                download_success, file_path, download_message = download_from_jiananguo()
                if download_success:
                    st.session_state.file_path = file_path
                    st.session_state.last_download = datetime.now()
                    st.session_state.schedule_data = None
                    st.session_state.refresh_counter += 1
                    st.session_state.force_current_time = True  # 下载后强制当前时段
                    st.success("班表已更新")
                else:
                    st.error(f"下载失败: {download_message}")
    
    st.markdown("---")  # 分割线
    
    # 显示更新日志（班表最后更新、下次自动刷新）
    info_text = []
    if st.session_state.last_download:
        info_text.append(f"班表最后更新: {st.session_state.last_download.strftime('%Y-%m-%d %H:%M:%S')}")
    if st.session_state.last_refresh:
        info_text.append(f"状态最后刷新: {st.session_state.last_refresh.strftime('%Y-%m-%d %H:%M:%S')}")
    # 计算下次自动刷新时间
    next_refresh = st.session_state.last_auto_refresh + timedelta(hours=1)
    info_text.append(f"下次自动刷新: {next_refresh.strftime('%H:%M:%S')}")
    
    if info_text:
        st.info(" | ".join(info_text))
    
    # 日期和时间选择栏（强制当前时段）
    col_date, col_time = st.columns(2)
    with col_date:
        # 强制选择当前日期
        default_date = datetime.now().date()
        view_date = st.date_input(
            "选择查看日期", 
            default_date, 
            key=f"date_{st.session_state.refresh_counter}"
        )
    
    with col_time:
        now = datetime.now()
        # 仅保留"当前时段"选项，禁止手动切换
        hour_options = ["当前时段"]
        default_idx = 0
        
        # 禁用选择器，防止手动修改
        selected_time_str = st.selectbox(
            "选择查看时间（已强制锁定当前时段）", 
            hour_options,
            index=default_idx,
            key=f"time_{st.session_state.refresh_counter}",
            disabled=True  # 禁用下拉选择
        )
        
        # 强制使用当前实时时间
        view_time = now.time()
    
    check_time = view_time
    
    # 处理T2夜班跨天问题：0-8点显示前一天的排班
    current_hour = check_time.hour
    if current_hour < 8:
        load_date = view_date - timedelta(days=1)
    else:
        load_date = view_date
    
    # 显示当前查看的时间和实际加载的排班日期
    weekdays = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]
    weekday = weekdays[view_date.weekday()]
    if current_hour < 8:
        load_weekday = weekdays[load_date.weekday()]
        st.info(f"当前查看时间: {view_date.strftime('%Y年%m月%d日')} {weekday} {check_time.strftime('%H:%M')} (显示{load_date.strftime('%Y年%m月%d日')} {load_weekday}的排班数据)")
    else:
        st.info(f"当前查看时间: {view_date.strftime('%Y年%m月%d日')} {weekday} {check_time.strftime('%H:%M')}")
    
    # 日期变更时清除旧数据缓存
    if st.session_state.last_load_date != load_date or st.session_state.force_current_time:
        st.session_state.schedule_data = None
        st.session_state.last_load_date = load_date
        st.session_state.force_current_time = False  # 重置标志（仅用于日期变更判断）
    
    # 加载并显示坐席数据
    with st.spinner("正在加载坐席数据，请稍候..."):
        if st.session_state.schedule_data is None:
            st.session_state.schedule_data = viewer.load_schedule_with_colors(st.session_state.file_path, load_date)
        schedule_df = st.session_state.schedule_data
    
    # 处理数据加载失败的情况
    if schedule_df is None or schedule_df.empty:
        st.error("未加载到有效坐席数据，请检查文件内容或日期匹配情况。")
        st.info("请点击顶部的「重新下载班表」按钮尝试更新数据")
        return
    
    # 按A/B/C席分类坐席数据
    categorized_data = viewer.categorize_by_seat(schedule_df, check_time)
    
    # 显示各席位在线人数统计
    st.subheader("📊 坐席统计")
    stats_cols = st.columns(3)
    for i, (seat, agents) in enumerate(categorized_data.items()):
        # 统计"搬砖中"（在线）的人数
        online_count = sum(1 for agent in agents if agent['status'] == '搬砖中')
        total_count = len(agents)
        seat_color = viewer.seat_colors[seat]
        
        with stats_cols[i]:
            st.markdown(create_stat_card(seat, online_count, total_count, seat_color), unsafe_allow_html=True)
    
    st.markdown("---")
    
    # 分栏显示各席位的坐席详情
    cols = st.columns(3)
    for i, (seat, agents) in enumerate(categorized_data.items()):
        with cols[i]:
            # 席位标题（带颜色背景）
            st.markdown(
                f"### <span style='background-color:{viewer.seat_colors[seat]}; color:black; padding:4px 8px; border-radius:4px; border: 1px solid #000000;'>{seat}</span>", 
                unsafe_allow_html=True
            )
            # 显示每个坐席的卡片
            if agents:
                for agent in agents:
                    st.markdown(create_agent_card(agent, viewer), unsafe_allow_html=True)
            else:
                st.info(f"当前无{seat}坐席值班")

if __name__ == "__main__":
    main()


