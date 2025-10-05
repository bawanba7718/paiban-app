import pandas as pd
import datetime
from datetime import datetime, time, timedelta, timezone
import streamlit as st
import openpyxl
from openpyxl import load_workbook
import os
import time as t
import tempfile
from webdav3.client import Client
import threading
import html

# 定义东八区时区（UTC+8）
TZ_UTC_8 = timezone(timedelta(hours=8))

class AgentViewer:
    def __init__(self):
        # 颜色-职位对应关系
        self.color_roles = {
            'FFC000': 'C席',
            'FFEE79': 'C席',
            'E2EFDA': 'C席',
            '91AADF': 'C席',
            'D9E1F2': 'C席',
            'EF949F': 'B席',
            'FADADE': 'B席',
            '8CDDFA': '休',
            'FFFF00': '休',
            'FFFFFF': 'A席',  # 无颜色/白色为A席
            'FEE796': 'C席',  # 新增FEE796颜色为C席
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
        
        # 班次时间定义（基础时间）
        self.shift_times = {
            'T1': {'start': time(8, 0), 'end': time(20, 0), 'name': '白班', 
                  'break_start': time(13, 0), 'break_end': time(14, 0)},
            'T2': {'start': time(20, 0), 'end': time(8, 0), 'name': '夜班',
                  'break_start': None, 'break_end': None},
            'M2': {'start': time(8, 0), 'end': time(17, 0), 'name': '早班',
                  'break_start': time(14, 0), 'break_end': time(15, 0)},  # 基础设置，会被A席规则覆盖
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

    def get_work_status(self, shift_code, seat, color_code, check_time=None):
        """
        修复A席M2休息时间判断逻辑，确保13:00-14:00显示"干饭中"
        新增FEE796颜色T1班次的特殊规则：8:00-17:00为C席，17:00-20:00为A席，休14:00-15:00
        """
        if not shift_code or str(shift_code).strip() == '':
            return "未排班", "#BFBFBF", seat
            
        shift_code = str(shift_code).strip()
        main_shift = None
        
        # 强制提取主班次（优先匹配完整班次代码）
        for s in sorted(self.shift_times.keys(), key=lambda x: len(x), reverse=True):
            if s in shift_code:
                main_shift = s
                break
                
        if not main_shift:
            return "未知班次", "#BFBFBF", seat
            
        # 复制基础班次时间
        shift = self.shift_times[main_shift].copy()
        
        # 重点修复：A席M2（无论颜色）统一休13:00-14:00
        # 扩大匹配范围，确保所有A席M2都应用此规则
        if seat == 'A席' and main_shift == 'M2':
            # 强制设置休息时间，覆盖任何基础设置
            shift['break_start'] = time(13, 0)
            shift['break_end'] = time(14, 0)
        
        # FEE796颜色T1班次特殊规则：8:00-17:00为C席，17:00-20:00为A席，休14:00-15:00
        elif color_code == 'FEE796' and main_shift == 'T1':
            shift['break_start'] = time(14, 0)
            shift['break_end'] = time(15, 0)
            
            # 动态调整席位：8:00-17:00为C席，17:00-20:00为A席
            check_time = check_time or datetime.now(TZ_UTC_8).time()
            if time(17, 0) <= check_time < time(20, 0):
                seat = 'A席'  # 17:00-20:00期间视为A席
            else:
                seat = 'C席'  # 8:00-17:00期间视为C席
        
        # EF949F T1 B席休14:00-15:00
        elif seat == 'B席' and color_code == 'EF949F' and main_shift == 'T1':
            shift['break_start'] = time(14, 0)
            shift['break_end'] = time(15, 0)
        
        # 其他规则保持不变
        elif seat == 'A席' and main_shift == 'D2':
            shift['break_start'] = time(14, 0)
            shift['break_end'] = time(15, 0)
        
        elif seat == 'C席' and color_code == 'FFC000' and main_shift == 'T1':
            shift['break_start'] = time(13, 0)
            shift['break_end'] = time(14, 0)
            
        elif seat == 'C席' and color_code == 'D9E1F2':
            shift['break_start'] = time(14, 0)
            shift['break_end'] = time(15, 0)
        
        elif seat == 'C席' and color_code == 'E2EFDA' and main_shift == 'T1':
            shift['break_start'] = time(14, 0)
            shift['break_end'] = time(15, 0)
        
        elif seat == 'C席' and color_code == 'E2EFDA' and main_shift == 'M2':
            shift['break_start'] = time(14, 0)
            shift['break_end'] = time(15, 0)
        
        elif seat == 'B席' and color_code == 'FADADE' and main_shift == 'M2':
            shift['break_start'] = time(13, 0)
            shift['break_end'] = time(14, 0)
            
        elif seat == 'A席' and main_shift == 'T1':
            shift['break_start'] = time(14, 0)
            shift['break_end'] = time(15, 0)
            
        # 使用东八区时间，默认当前时间
        check_time = check_time or datetime.now(TZ_UTC_8).time()
        
        # 解构时间参数
        start, end = shift['start'], shift['end']
        break_start, break_end = shift.get('break_start'), shift.get('break_end')
        
        # 判断是否在工作时间内（优化跨天逻辑）
        is_night_shift = main_shift == 'T2'
        in_work_time = False
        
        if is_night_shift:
            # 夜班：20:00-次日08:00
            in_work_time = (check_time >= start) or (check_time < end)
        else:
            # 白班/早班：正常时间范围
            in_work_time = start <= check_time < end
            
        # 判断是否在上班路上
        is_on_the_way = False
        if not in_work_time and not is_night_shift:
            is_on_the_way = check_time < start
            
        # 重点修复：休息时间判断逻辑，确保13:00-14:00被正确识别
        in_break_time = False
        if break_start and break_end and in_work_time:
            # 确保休息时间是当天范围内（非跨天）
            if break_start < break_end:
                # 正常时间范围（如13:00-14:00）
                in_break_time = break_start <= check_time < break_end
            else:
                # 跨天休息时间（本系统不适用，但保留逻辑）
                in_break_time = check_time >= break_start or check_time < break_end
        
        # 确定最终状态
        if is_on_the_way:
            return "正在路上", "#BFBFBF", seat
        elif not in_work_time:
            return "已回家", "#BFBFBF", seat
        elif in_break_time:
            return "干饭中", "orange", seat
        else:
            return "搬砖中", "green", seat

    def get_cell_color(self, cell):
        try:
            if cell and cell.fill and cell.fill.start_color:
                color = cell.fill.start_color.rgb
                if color:
                    color_str = str(color).upper()
                    if color_str.startswith('FF'):
                        color_str = color_str[2:]  # 去除alpha通道
                    elif len(color_str) == 8:
                        color_str = color_str[2:]
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
            
            for idx, col in enumerate(df_main.columns):
                col_str = str(col)
                if (target_date_str in col_str or 
                    target_date.strftime('%m-%d') in col_str or
                    target_date.strftime('%Y/%m/%d') in col_str or
                    target_date.strftime('%m/%d') in col_str):
                    today_col_idx = idx
                    break
            
            if today_col_idx is None:
                st.warning(f"未找到 {target_date_str} 的排班列，可能该日期无排班数据")
                return pd.DataFrame()
            
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
                    
                    if (not shift_code or 
                        shift_code.strip() in ['', '休', '休息']):
                        continue
                    
                    seat = self.color_roles.get(color_code, 'A席')
                    
                    person_info = {
                        'name': name,
                        'id': str(row[2].value).strip() if row[2].value else '',
                        'workplace': str(row[0].value).strip() if row[0].value else '',
                        'shift': shift_code,
                        'color': color_code,
                        'seat': seat,
                        'status': '',
                        'status_color': '',
                        'actual_seat': seat,  # 新增字段，用于存储实际席位
                        'date': target_date  # 添加日期信息，用于区分班次
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
            status, status_color, actual_seat = self.get_work_status(
                person['shift'], 
                person['seat'], 
                person['color'],
                check_time
            )
            person['status'] = status
            person['status_color'] = status_color
            person['actual_seat'] = actual_seat  # 存储实际席位（考虑动态变化）
            
            seat = actual_seat  # 使用实际席位进行分类
            if seat in result:
                result[seat].append(person)
            else:
                result['A席'].append(person)
        
        # 排序逻辑：优先状态（搬砖中在前），同状态按上班时间排序
        status_priority = {
            '搬砖中': 3,
            '干饭中': 2,
            '正在路上': 1,
            '已回家': 0,
            '未排班': -1,
            '未知班次': -2
        }
        
        for cat in result:
            result[cat].sort(key=lambda x: (
                -status_priority.get(x['status'], -3),
                self.get_shift_start_time(x['shift'])
            ))
        
        return result
    
    def get_shift_start_time(self, shift_code):
        if not shift_code or str(shift_code).strip() == '':
            return time(23, 59, 59)
            
        shift_code = str(shift_code).strip()
        
        for s in self.shift_times:
            if s in shift_code:
                return self.shift_times[s]['start']
        
        return time(23, 59, 59)

def download_from_jiananguo():
    try:
        jiananguo_email = st.secrets.get("JIANANGUO_EMAIL", "hanyong@foxmail.com")
        jiananguo_password = st.secrets.get("JIANANGUO_PASSWORD", "ah5fb6yahy62b8rt")
        
        options = {
            'webdav_hostname': 'https://dav.jianguoyun.com/dav/',
            'webdav_login': jiananguo_email,
            'webdav_password': jiananguo_password
        }
        
        client = Client(options)
        remote_file = '我的坚果云/排班.xlsx'
        
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            local_file = tmp_file.name
        
        client.download_sync(remote_path=remote_file, local_path=local_file)
        
        if os.path.exists(local_file) and os.path.getsize(local_file) > 0:
            return True, local_file, ""
        else:
            return False, None, "从坚果云下载文件失败"
            
    except Exception as e:
        return False, None, f"下载失败: {str(e)}"

def create_compact_agent_card(person_info, viewer):
    """创建紧凑型坐席卡片"""
    status_icon = viewer.status_icons.get(person_info['status'], '❓')
    
    # 状态颜色
    if person_info['status'] in ["正在路上", "已回家"]:
        status_color = "#BFBFBF"
        bg_color = "#F5F5F5"  # 浅灰色背景
        border_color = "#DDD"
    else:
        status_color = person_info['status_color']
        seat_type = person_info.get('actual_seat', person_info['seat'])
        bg_color = f"#{person_info['color']}" if seat_type in ['B席', 'C席'] else "#FFFFFF"
        border_color = "#333"
    
    # 显示实际席位
    display_seat = person_info.get('actual_seat', person_info['seat'])
    
    # 紧凑卡片设计
    card_html = f"""
    <div style="background-color: {bg_color}; border: 1px solid {border_color}; border-radius: 4px; padding: 6px; margin: 2px; min-height: 60px; display: flex; flex-direction: column; justify-content: center;">
        <div style="font-size: 14px; font-weight: bold; text-align: center; margin-bottom: 4px;">{person_info['name']}</div>
        <div style="display: flex; justify-content: space-between; align-items: center;">
            <div style="font-size: 12px; color: #666;">{person_info['workplace']}</div>
            <div style="font-size: 16px;">{status_icon}</div>
        </div>
        <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 2px;">
            <div style="font-size: 11px; font-weight: bold;">{person_info['shift']}</div>
            <div style="font-size: 11px; color: {status_color}; font-weight: bold;">{person_info['status']}</div>
        </div>
    </div>
    """
    
    return card_html

def update_current_time():
    weekdays = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]
    now = datetime.now(TZ_UTC_8)
    weekday = weekdays[now.weekday()]
    return now.strftime(f"%Y年%m月%d日 {weekday} %H:%M:%S")

def auto_refresh_time(placeholder):
    while True:
        if not st.session_state.get('auto_refresh', True):
            t.sleep(1)
            continue
        placeholder.markdown(f"### 当前时间: {update_current_time()}")
        
        current_minute = datetime.now(TZ_UTC_8).minute
        if current_minute == 0 and not st.session_state.get('hour_refresh_done', False):
            st.session_state.hour_refresh_done = True
            st.session_state.refresh_counter += 1
            st.session_state.schedule_data = {}
            st.rerun()
        elif current_minute != 0:
            st.session_state.hour_refresh_done = False
            
        t.sleep(1)

def filter_data_by_workplace(df, workplace):
    """根据职场筛选数据"""
    if workplace == "全部":
        return df
    elif workplace in ["重庆", "北京"]:
        return df[df['workplace'] == workplace]
    else:
        return df

def filter_data_by_name(df, name_query):
    """根据姓名查询筛选数据"""
    if not name_query:
        return df
    return df[df['name'].str.contains(name_query, case=False, na=False)]

def main():
    st.set_page_config(
        page_title="综合组在线坐席", 
        layout="wide",
        page_icon="📊"
    )
    
    # 初始化session state
    if 'file_path' not in st.session_state:
        st.session_state.file_path = None
    if 'last_download' not in st.session_state:
        st.session_state.last_download = None
    if 'last_refresh' not in st.session_state:
        st.session_state.last_refresh = None
    if 'schedule_data' not in st.session_state:
        st.session_state.schedule_data = {}
    if 'auto_refresh' not in st.session_state:
        st.session_state.auto_refresh = True
    if 'refresh_counter' not in st.session_state:
        st.session_state.refresh_counter = 0
    if 'hour_refresh_done' not in st.session_state:
        st.session_state.hour_refresh_done = False
    if 'workplace_filter' not in st.session_state:
        st.session_state.workplace_filter = "全部"
    if 'name_query' not in st.session_state:
        st.session_state.name_query = ""
    
    # 初始化查看器
    viewer = AgentViewer()
    
    # 首次运行或文件不存在时下载排班文件
    if st.session_state.file_path is None or not os.path.exists(st.session_state.file_path):
        with st.spinner("正在加载排班文件..."):
            download_success, file_path, download_message = download_from_jiananguo()
            if download_success:
                st.session_state.file_path = file_path
                st.session_state.last_download = datetime.now(TZ_UTC_8)
            else:
                st.error(f"加载失败: {download_message}")
                st.stop()
    
    # 主界面 - 看板式布局
    # 顶部标题栏
    col_logo, col_title, col_date = st.columns([1, 3, 2])
    
    with col_logo:
        # 显示HealthLink远盟康健logo - 绿色
        logo_html = """
        <div style="display: flex; align-items: center; justify-content: center; padding: 5px;">
            <div style="text-align: center;">
                <h2 style="margin: 0; color: #2E8B57; font-weight: bold;">HealthLink</h2>
                <p style="margin: 0; color: #2E8B57; font-size: 12px;">远盟康健®</p>
            </div>
        </div>
        """
        st.markdown(logo_html, unsafe_allow_html=True)
    
    with col_title:
        st.title("综合组在线坐席")
    
    with col_date:
        current_datetime = st.empty()
        if 'time_thread' not in st.session_state:
            st.session_state.time_thread = threading.Thread(
                target=auto_refresh_time, 
                args=(current_datetime,), 
                daemon=True
            )
            st.session_state.time_thread.start()
    
    # 控制按钮栏
    col_controls = st.columns([2, 1, 1, 1])
    
    with col_controls[0]:
        # 搜索和筛选区域
        col_search, col_filter = st.columns([2, 1])
        with col_search:
            name_query = st.text_input(
                "搜索姓名",
                placeholder="输入姓名关键字...",
                key=f"name_query_{st.session_state.refresh_counter}"
            )
            st.session_state.name_query = name_query
        
        with col_filter:
            workplace_filter = st.selectbox(
                "选择职场",
                ["全部", "重庆", "北京"],
                key=f"workplace_{st.session_state.refresh_counter}"
            )
            st.session_state.workplace_filter = workplace_filter
    
    with col_controls[1]:
        view_date = st.date_input(
            "选择日期", 
            datetime.now(TZ_UTC_8).date(),
            key=f"date_{st.session_state.refresh_counter}"
        )
    
    with col_controls[2]:
        hour_options = [f"{h:02d}:00" for h in range(24)]
        current_hour_str = f"{datetime.now(TZ_UTC_8).hour:02d}:00"
        
        default_idx = hour_options.index(current_hour_str) if current_hour_str in hour_options else 0
        
        selected_time_str = st.selectbox(
            "选择时间", 
            hour_options,
            index=default_idx,
            key=f"time_{st.session_state.refresh_counter}"
        )
        
        hour = int(selected_time_str.split(":")[0])
        view_time = time(hour, 0)
    
    with col_controls[3]:
        col_refresh1, col_refresh2 = st.columns(2)
        with col_refresh1:
            if st.button("🔄 刷新", use_container_width=True):
                st.session_state.last_refresh = datetime.now(TZ_UTC_8)
                st.session_state.refresh_counter += 1
                st.session_state.schedule_data = {}
                st.success("状态已刷新")
        
        with col_refresh2:
            if st.button("📥 重载", use_container_width=True):
                with st.spinner("重新加载中..."):
                    download_success, file_path, download_message = download_from_jiananguo()
                    if download_success:
                        st.session_state.file_path = file_path
                        st.session_state.last_download = datetime.now(TZ_UTC_8)
                        st.session_state.schedule_data = {}
                        st.session_state.refresh_counter += 1
                        st.success("数据已更新")
                    else:
                        st.error(f"加载失败: {download_message}")
    
    st.markdown("---")
    
    # 显示最后更新时间
    info_text = []
    if st.session_state.last_download:
        info_text.append(f"班表更新: {st.session_state.last_download.strftime('%H:%M:%S')}")
    if st.session_state.last_refresh:
        info_text.append(f"状态刷新: {st.session_state.last_refresh.strftime('%H:%M:%S')}")
    
    # 显示下次自动刷新时间
    now = datetime.now(TZ_UTC_8)
    next_hour = (now.replace(minute=0, second=0, microsecond=0) + timedelta(hours=1))
    info_text.append(f"下次刷新: {next_hour.strftime('%H:%M:%S')}")
    
    if info_text:
        st.caption(" | ".join(info_text))
    
    # 显示当前查看时间
    weekdays = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]
    weekday = weekdays[view_date.weekday()]
    
    current_hour = view_time.hour
    
    # 修复7-8点时间段班次显示问题
    # 在7-8点时间段，需要同时显示昨日的T2班次和今日的早班班次
    if 7 <= current_hour < 8:
        # 同时加载昨日和今日的数据
        yesterday_date = view_date - timedelta(days=1)
        
        # 使用日期字符串作为缓存键
        yesterday_key = yesterday_date.strftime('%Y-%m-%d')
        today_key = view_date.strftime('%Y-%m-%d')
        
        # 加载昨日数据
        if yesterday_key not in st.session_state.schedule_data:
            with st.spinner(f"正在加载{yesterday_date.strftime('%Y年%m月%d日')}的坐席数据，请稍候..."):
                st.session_state.schedule_data[yesterday_key] = viewer.load_schedule_with_colors(
                    st.session_state.file_path, 
                    yesterday_date
                )
        
        # 加载今日数据
        if today_key not in st.session_state.schedule_data:
            with st.spinner(f"正在加载{view_date.strftime('%Y年%m月%d日')}的坐席数据，请稍候..."):
                st.session_state.schedule_data[today_key] = viewer.load_schedule_with_colors(
                    st.session_state.file_path, 
                    view_date
                )
        
        yesterday_df = st.session_state.schedule_data[yesterday_key]
        today_df = st.session_state.schedule_data[today_key]
        
        # 合并数据：只保留昨日的T2班次和今日的非T2班次
        combined_data = []
        
        # 添加昨日的T2班次
        if yesterday_df is not None and not yesterday_df.empty:
            for _, person in yesterday_df.iterrows():
                if 'T2' in str(person['shift']):
                    person_copy = person.copy()
                    person_copy['is_yesterday'] = True
                    person_copy['date'] = yesterday_date
                    combined_data.append(person_copy)
        
        # 添加今日的非T2班次
        if today_df is not None and not today_df.empty:
            for _, person in today_df.iterrows():
                if 'T2' not in str(person['shift']):
                    person_copy = person.copy()
                    person_copy['is_yesterday'] = False
                    person_copy['date'] = view_date
                    combined_data.append(person_copy)
        
        schedule_df = pd.DataFrame(combined_data) if combined_data else pd.DataFrame()
        
        st.info(f"当前查看: {view_date.strftime('%Y年%m月%d日')} {weekday} {view_time.strftime('%H:%M')} (显示{yesterday_date.strftime('%m-%d')}的T2班次和{view_date.strftime('%m-%d')}的其他班次)")
    
    elif current_hour < 7:
        # 7点之前显示前一天的排班
        load_date = view_date - timedelta(days=1)
        load_weekday = weekdays[load_date.weekday()]
        st.info(f"当前查看: {view_date.strftime('%Y年%m月%d日')} {weekday} {view_time.strftime('%H:%M')} (显示{load_date.strftime('%Y年%m月%d日')} {load_weekday}的排班数据)")
        
        # 使用日期字符串作为缓存键
        load_date_key = load_date.strftime('%Y-%m-%d')
        
        # 加载对应日期的数据
        if load_date_key not in st.session_state.schedule_data:
            with st.spinner(f"正在加载{load_date.strftime('%Y年%m月%d日')}的坐席数据，请稍候..."):
                st.session_state.schedule_data[load_date_key] = viewer.load_schedule_with_colors(
                    st.session_state.file_path, 
                    load_date
                )
        
        schedule_df = st.session_state.schedule_data[load_date_key]
    
    else:
        # 8点及之后显示当天的排班
        load_date = view_date
        st.info(f"当前查看: {view_date.strftime('%Y年%m月%d日')} {weekday} {view_time.strftime('%H:%M')}")
        
        # 使用日期字符串作为缓存键
        load_date_key = load_date.strftime('%Y-%m-%d')
        
        # 加载对应日期的数据
        if load_date_key not in st.session_state.schedule_data:
            with st.spinner(f"正在加载{load_date.strftime('%Y年%m月%d日')}的坐席数据，请稍候..."):
                st.session_state.schedule_data[load_date_key] = viewer.load_schedule_with_colors(
                    st.session_state.file_path, 
                    load_date
                )
        
        schedule_df = st.session_state.schedule_data[load_date_key]
    
    if schedule_df is None or schedule_df.empty:
        st.warning(f"未找到有效坐席数据")
        return
    
    # 应用职场筛选
    schedule_df = filter_data_by_workplace(schedule_df, st.session_state.workplace_filter)
    
    # 应用姓名查询
    schedule_df = filter_data_by_name(schedule_df, st.session_state.name_query)
    
    if schedule_df.empty:
        st.warning(f"未找到符合条件的坐席数据")
        return
    
    # 按A/B/C席分类显示坐席
    categorized_data = viewer.categorize_by_seat(schedule_df, view_time)
    
    # 看板式布局 - 三列并排
    st.subheader(f"{view_date.strftime('%Y年%m月%d日')} {weekday} 坐席看板")
    
    # 创建三列
    col_a, col_b, col_c = st.columns(3)
    
    # A席看板
    with col_a:
        agents_a = categorized_data.get('A席', [])
        online_count_a = sum(1 for agent in agents_a if agent['status'] == '搬砖中')
        total_count_a = len(agents_a)
        
        # 席位标题
        st.markdown(f"""
        <div style="background-color: #FFFFFF; border: 2px solid #333; border-radius: 6px; padding: 8px; margin-bottom: 8px; text-align: center;">
            <h3 style="margin: 0; color: #333;">A席 ({online_count_a}/{total_count_a})</h3>
        </div>
        """, unsafe_allow_html=True)
        
        # 坐席网格
        if agents_a:
            # 计算每行显示的坐席数量
            cols_per_row = 2
            for i in range(0, len(agents_a), cols_per_row):
                cols = st.columns(cols_per_row)
                for j in range(cols_per_row):
                    if i + j < len(agents_a):
                        with cols[j]:
                            st.markdown(create_compact_agent_card(agents_a[i + j], viewer), unsafe_allow_html=True)
        else:
            st.info("暂无A席坐席")
    
    # B席看板
    with col_b:
        agents_b = categorized_data.get('B席', [])
        online_count_b = sum(1 for agent in agents_b if agent['status'] == '搬砖中')
        total_count_b = len(agents_b)
        
        # 席位标题
        st.markdown(f"""
        <div style="background-color: #EF949F; border: 2px solid #333; border-radius: 6px; padding: 8px; margin-bottom: 8px; text-align: center;">
            <h3 style="margin: 0; color: #333;">B席 ({online_count_b}/{total_count_b})</h3>
        </div>
        """, unsafe_allow_html=True)
        
        # 坐席网格
        if agents_b:
            # 计算每行显示的坐席数量
            cols_per_row = 2
            for i in range(0, len(agents_b), cols_per_row):
                cols = st.columns(cols_per_row)
                for j in range(cols_per_row):
                    if i + j < len(agents_b):
                        with cols[j]:
                            st.markdown(create_compact_agent_card(agents_b[i + j], viewer), unsafe_allow_html=True)
        else:
            st.info("暂无B席坐席")
    
    # C席看板
    with col_c:
        agents_c = categorized_data.get('C席', [])
        online_count_c = sum(1 for agent in agents_c if agent['status'] == '搬砖中')
        total_count_c = len(agents_c)
        
        # 席位标题
        st.markdown(f"""
        <div style="background-color: #FFC000; border: 2px solid #333; border-radius: 6px; padding: 8px; margin-bottom: 8px; text-align: center;">
            <h3 style="margin: 0; color: #333;">C席 ({online_count_c}/{total_count_c})</h3>
        </div>
        """, unsafe_allow_html=True)
        
        # 坐席网格
        if agents_c:
            # 计算每行显示的坐席数量
            cols_per_row = 2
            for i in range(0, len(agents_c), cols_per_row):
                cols = st.columns(cols_per_row)
                for j in range(cols_per_row):
                    if i + j < len(agents_c):
                        with cols[j]:
                            st.markdown(create_compact_agent_card(agents_c[i + j], viewer), unsafe_allow_html=True)
        else:
            st.info("暂无C席坐席")

if __name__ == "__main__":
    main()
