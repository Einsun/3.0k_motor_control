"""
$3.0k_motor_control-ver3.3b
2025/3/13

192.168.1.1 网关
192.168.1.121:8802 3.6w电机
192.168.1.122:8802 3k电机
"""
import csv
import os
import json
import struct
import sys
import time
import socket
import threading
from datetime import datetime

from openpyxl import Workbook
from PyQt5 import uic, QtGui
from PyQt5.Qt import QThread, pyqtSignal, Qt
from PyQt5.QtWidgets import QWidget, QApplication, QMessageBox, QLCDNumber

current_file_path = __file__
file_name = os.path.basename(current_file_path)[:-3]
print("当前版本:", file_name)


class MotorController(QWidget):
    def __init__(self):
        super().__init__()
        self.config = MotorConfig()  # 先初始化config

        self._init_ui()
        self._init_variables()
        self._setup_connections()
        self.sock = None  # 唯一的socket连接
        self.thread = None
        self.crc_helper = CRCHelper()
        # self.data_buffer = []  # 数据采集缓冲区

    def _init_ui(self):
        self.ui = uic.loadUi("./_internal/motor_control.ui")
        self.ui.labver.setText(file_name)
        self._setup_ui_defaults()

    def _setup_ui_defaults(self):
        """设置UI元素的默认状态"""
        self.ui.labaddr.setText(f"{self.config.ip_address}:{self.config.port}")
        self.ui.introt.setValidator(QtGui.QIntValidator(0, self.config.max_speed))
        self.ui.intupt.setValidator(QtGui.QIntValidator(0, 6500))
        self.ui.intdot.setValidator(QtGui.QIntValidator(0, 6500))

        self.ui.labelCH0.setText(self.config.modbus_head[0])
        self.ui.labelCH1.setText(self.config.modbus_head[1])
        self.ui.labelCH2.setText(self.config.modbus_head[2])
        self.ui.labelCH3.setText(self.config.modbus_head[3])

        self.ui.labelCH4.setText(self.config.modbus_head[4])
        self.ui.labelCH5.setText(self.config.modbus_head[5])
        self.ui.labelCH6.setText(self.config.modbus_head[6])
        self.ui.labelCH7.setText(self.config.modbus_head[7])

        self._enable_controls(False)

    # 初始化变量
    def _init_variables(self):
        """初始化变量"""
        self.csv_file = None  # CSV文件对象
        self.csv_writer = None  # CSV写入器
        self.csv_filename = ""  # 当前CSV文件名
        self.motor_params = {
            'max_speed': self.config.max_speed,
            'rotation_ratio': self.config.rotation_ratio,
            'sample_interval': self.config.sample_interval,
            'is_running': 0
        }
        self.current_state = {
            'speed': 0,
            'torque': 0.0,
            'voltage': 0,
            'current': 0.0,
            'power': 0.0,
            'run_status': 0,
            'torque_meter_torque': 0.0,  # 转矩仪转矩
            'torque_meter_speed': 0.0,  # 转矩仪转速
            'torque_meter_power': 0.0  # 转矩仪功率
        }
        self.run_status_text = ['未知', '运行中', '运行中(反向)', '停止']

    def _setup_connections(self):
        """连接所有信号和槽"""
        connections = [
            (self.ui.btnlink.clicked, self._handle_connection),
            (self.ui.btnread.clicked, self.read_motor_status),
            (self.ui.btnupt.clicked, self.set_acceleration_time),
            (self.ui.btndot.clicked, self.set_deceleration_time),
            (self.ui.btnrot.clicked, self.set_rotation_speed),
            (self.ui.btnctrl1.clicked, self.set_local_control),
            (self.ui.btnctrl2.clicked, self.set_remote_control),
            (self.ui.rbtnforward.clicked, self.set_forward_rotation),
            (self.ui.rbtnreserve.clicked, self.set_reverse_rotation),
            (self.ui.btnrun.clicked, self.start_motor),
            (self.ui.btnst1.clicked, self.stop_motor_soft),
            (self.ui.btnst2.clicked, self.stop_motor_hard),
            (self.ui.pbtndaq.clicked, self.toggle_data_collection),
            (self.ui.btngra.clicked, self.send_custom_command),
            (self.ui.btncln.clicked, self.clear_command_display)
        ]
        for signal, slot in connections:
            signal.connect(slot)

    def _ensure_connection(self):
        """确保socket连接有效"""
        if self.sock is None:
            self.log_message('error', '未建立连接')
            return False
        return True

    def _enable_controls(self, enabled):
        """统一启用/禁用控件"""
        controls = [
            self.ui.Lcontrol, self.ui.btnctrl1, self.ui.btnctrl2,
            self.ui.btngra, self.ui.btnread, self.ui.btnrun
        ]
        for control in controls:
            control.setEnabled(enabled)

    # 处理连接 / 断开连接
    def _handle_connection(self):
        """处理连接/断开连接"""
        if self.sock is None:
            self._connect_to_motor()
        else:
            self._disconnect_from_motor()

    # 建立socket连接
    def _connect_to_motor(self):
        """建立socket连接"""
        try:
            self.sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.sock.setsockopt(socket.IPPROTO_TCP, socket.TCP_NODELAY, 1)  # 关闭小包合并算法
            self.sock.settimeout(0.05)
            self.sock.connect((self.config.ip_address, self.config.port))

            self._enable_controls(True)
            self.ui.btnlink.setText("断开连接")
            self.ui.labconn.setText("已连接")
            self.log_message('info', '连接成功')
            return True
        except Exception as e:
            self.log_message('error', f'连接失败: {str(e)}')
            self._close_socket()
            return False

    # 断开socket连接
    def _disconnect_from_motor(self):
        """断开socket连接"""
        self._close_socket()
        self._enable_controls(False)
        self.ui.btnlink.setText("连接")
        self.ui.labconn.setText("未连接")
        self.log_message('info', '连接已关闭')

    # 安全关闭socket
    def _close_socket(self):
        """安全关闭socket"""
        if self.sock:
            try:
                self.sock.close()
            except:
                pass
            finally:
                self.sock = None

    # 发送命令并返回响应
    def send_command(self, command, expected_response_prefix='', retries=3):
        """发送命令并返回响应"""
        if not self._ensure_connection():
            return False, None

        for attempt in range(retries):
            try:
                # 添加CRC并发送
                data_with_crc = self.crc_helper.add_crc(bytes.fromhex(command))
                self.sock.sendall(data_with_crc)

                # 接收响应
                response = self.sock.recv(1024)
                if not response:
                    continue

                # 验证CRC
                crc_valid, payload = self.crc_helper.verify_crc(response)
                if not crc_valid:
                    self.log_message('warning', 'CRC校验失败')
                    continue

                hex_response = payload.hex()
                if expected_response_prefix and not hex_response.startswith(expected_response_prefix):
                    self.log_message('warning', f'响应前缀不匹配: 期望 {expected_response_prefix}, 收到 {hex_response[:4]}')
                    continue

                return True, hex_response
            except socket.timeout:
                self.log_message('warning', f'命令超时 (尝试 {attempt + 1}/{retries})')
            except Exception as e:
                self.log_message('error', f'发送命令出错: {str(e)}')
                self._close_socket()
                break

        return False, None

    # 读取电机状态
    def read_motor_status(self):
        """读取电机状态"""
        commands = [
            ('010370000007', '0103', self._parse_motor_parameters),
            ('0103F0110002', '0103', self._parse_timing_parameters),
            ('010330000001', '0103', self._parse_run_status),
            ('020300000006', '0203', self._parse_torque_meter)  # 新增转矩仪数据读取
        ]

        for cmd, prefix, parser in commands:
            success, response = self.send_command(cmd, prefix)
            if success:
                parser(response)
            else:
                self.log_message('error', f'读取 {cmd} 失败')

    # 解析电机参数响应
    def _parse_motor_parameters(self, response):
        """解析电机参数响应"""
        try:
            read_data = struct.unpack('>hhhhhhh', bytes.fromhex(response)[3:17])
            self.current_state['speed'] = round(read_data[0] * 0.6 * self.motor_params['rotation_ratio'])
            self.current_state['set_speed'] = round(read_data[1] * 0.6 * self.motor_params['rotation_ratio'])
            self.current_state['voltage'] = read_data[3]
            self.current_state['current'] = read_data[4] / 100
            self.current_state['power'] = read_data[5] / 10
            self.current_state['torque'] = read_data[6] / 10

            # 更新UI
            self.ui.ledoutrot.display(self.current_state['speed'])
            self.ui.ledsetrot.display(self.current_state['set_speed'])
            self.ui.ledoutvot.display(self.current_state['voltage'])
            self.ui.ledoutcur.display(self.current_state['current'])
            self.ui.ledoutpow.display(self.current_state['power'])
            self.ui.ledouttor.display(self.current_state['torque'])
            self.ui.introt.setText(str(self.current_state['set_speed']))
        except Exception as e:
            self.log_message('error', f'解析电机参数出错: {str(e)}')

    # 解析时间参数响应
    def _parse_timing_parameters(self, response):
        """解析时间参数响应"""
        try:
            read_data = struct.unpack('>hh', bytes.fromhex(response)[3:7])
            upt = read_data[0]/10
            dot = read_data[1]/10

            self.ui.ledupt.display(upt)
            self.ui.leddot.display(dot)
            self.ui.intupt.setText(str(upt))
            self.ui.intdot.setText(str(dot))
        except Exception as e:
            self.log_message('error', f'解析时间参数出错: {str(e)}')

    # 解析运行状态响应
    def _parse_run_status(self, response):
        """解析运行状态响应"""
        try:
            read_data = struct.unpack('>h', bytes.fromhex(response)[3:5])
            self.motor_params['is_running'] = read_data[0]
            status = self.motor_params['is_running']

            if 0 <= status < len(self.run_status_text):
                self.ui.labisrun.setText(self.run_status_text[status])
            else:
                self.ui.labisrun.setText("未知状态")
        except Exception as e:
            self.log_message('error', f'解析运行状态出错: {str(e)}')

    # 解析转矩仪数据响应
    def _parse_torque_meter(self, response):
        """解析转矩仪数据响应"""
        try:
            read_data = struct.unpack('>iii', bytes.fromhex(response)[3:15])
            self.current_state['torque_meter_torque'] = read_data[0] / 100
            self.current_state['torque_meter_speed'] = read_data[1] / 10
            self.current_state['torque_meter_power'] = read_data[2] / 100

            # 更新转矩仪显示
            self.ui.ledreadtor.display(self.current_state['torque_meter_torque'])
            self.ui.ledreadrot.display(self.current_state['torque_meter_speed'])
            self.ui.ledreadpwr.display(self.current_state['torque_meter_power'])
        except Exception as e:
            self.log_message('error', f'解析转矩仪数据出错: {str(e)}')

    # 设置加速时间
    def set_acceleration_time(self):
        """设置加速时间"""
        upt = self.ui.intupt.text()
        if upt and 0 < int(upt) < 6500:
            cmd = f'0106F011{int(upt) * 10:04X}'
            success, _ = self.send_command(cmd, '0106')
            if success:
                self.ui.ledupt.display(upt)
                self.log_message('info', f'加速时间设置为 {upt} 秒')
        else:
            self.log_message('error', '加速时间必须在0-6500之间')

    # 设置减速时间
    def set_deceleration_time(self):
        """设置减速时间"""
        dot = self.ui.intdot.text()
        if dot and 0 < int(dot) < 6500:
            cmd = f'0106F012{int(dot) * 10:04X}'
            success, _ = self.send_command(cmd, '0106')
            if success:
                self.ui.leddot.display(dot)
                self.log_message('info', f'减速时间设置为 {dot} 秒')
        else:
            self.log_message('error', '减速时间必须在0-6500之间')

    # 设置转速
    def set_rotation_speed(self):
        """设置转速"""
        speed = self.ui.introt.text()
        if speed and 0 <= int(speed) <= self.motor_params['max_speed']:
            value = round(int(speed) / (self.motor_params['rotation_ratio'] * 0.3))
            cmd = f'01061000{value:04X}'
            success, _ = self.send_command(cmd, '0106')
            if success:
                self.ui.ledsetrot.display(speed)
                self.log_message('info', f'转速设置为 {speed} RPM')
        else:
            self.log_message('error', f'转速必须在0-{self.motor_params["max_speed"]}之间')

    # 设置正转
    def set_forward_rotation(self):
        """设置正转"""
        cmd = '0106F0090000'
        success, _ = self.send_command(cmd, '0106')
        if success:
            self.ui.btnrun.setEnabled(True)
            self.log_message('info', '旋向设置为正转')

    # 设置反转
    def set_reverse_rotation(self):
        """设置反转"""
        cmd = '0106F0090001'
        success, _ = self.send_command(cmd, '0106')
        if success:
            self.ui.btnrun.setEnabled(True)
            self.log_message('info', '旋向设置为反转')

    # 设置为本地控制
    def set_local_control(self):
        """设置为本地控制"""
        commands = [
            ('0106F0020000', '0106'),
            ('0106F0030004', '0106')
        ]
        for cmd, prefix in commands:
            success, _ = self.send_command(cmd, prefix)
            if not success:
                return
        self.log_message('info', '设置为面板操作')

    # 设置为远程控制
    def set_remote_control(self):
        """设置为远程控制"""
        commands = [
            ('0106F0020002', '0106'),
            ('0106F0030009', '0106')
        ]
        for cmd, prefix in commands:
            success, _ = self.send_command(cmd, prefix)
            if not success:
                return
        self.log_message('info', '设置为远程通讯操作')

    # 启动电机
    def start_motor(self):
        """启动电机"""
        cmd = '010620000001'
        success, _ = self.send_command(cmd, '0106')
        if success:
            self.log_message('info', '电机启动')

    # 软停止电机
    def stop_motor_soft(self):
        """软停止电机"""
        cmd = '010620000006'  # 减速停机
        success, _ = self.send_command(cmd, '0106')
        if success:
            self.log_message('info', '电机减速停止')

    # 急停电机
    def stop_motor_hard(self):
        """急停电机"""
        cmd = '010620000005'  # 自由停机
        success, _ = self.send_command(cmd, '0106')
        if success:
            self.log_message('info', '电机急停')

    # 切换数据采集状态
    def toggle_data_collection(self):
        """切换数据采集状态"""
        if self.thread and self.thread.isRunning():
            self._stop_data_collection()
        else:
            self._start_data_collection()

    # 启动数据采集线程
    def _start_data_collection(self):
        """启动数据采集线程"""
        if not self._ensure_connection():
            return

        self.thread = DataCollectionThread(
            interval=self.motor_params['sample_interval'],
            controller=self
        )
        self.thread.data_ready.connect(self.update_data_display)
        self.thread.start()

        self.ui.pbtndaq.setText("停止采集")
        self.ui.cboxdaq.setEnabled(False)
        self.ui.btnlink.setEnabled(False)
        self.ui.btnread.setEnabled(False)
        self.ui.btnctrl1.setEnabled(False)
        self.ui.btnctrl2.setEnabled(False)
        self.ui.btnupt.setEnabled(False)
        self.ui.btndot.setEnabled(False)
        self.log_message('info', '数据采集已启动')

        if self.ui.cboxdaq.isChecked():
            self._start_csv_writer()

    # 停止数据采集线程
    def _stop_data_collection(self):
        """停止数据采集线程"""
        if self.thread:
            self.thread.stop()
            self.thread.wait(2000)

            if self.ui.cboxdaq.isChecked():
                # 关闭CSV写入器
                self._close_csv_writer()
                # self._save_collected_data()

        self.ui.pbtndaq.setText("开始采集")
        self.ui.cboxdaq.setEnabled(True)
        self.ui.btnlink.setEnabled(True)
        self.ui.btnread.setEnabled(True)
        self.ui.btnctrl1.setEnabled(True)
        self.ui.btnctrl2.setEnabled(True)
        self.ui.btnupt.setEnabled(True)
        self.ui.btndot.setEnabled(True)
        self.log_message('info', '数据采集已停止')

    # 启动CSV写入器
    def _start_csv_writer(self):
        """启动CSV写入器"""
        if not self.ui.cboxdaq.isChecked():
            return

        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            self.csv_filename = f"motor_data_{timestamp}.csv"

            # 创建CSV文件并写入表头
            self.csv_file = open(self.csv_filename, 'w', newline='', encoding='utf-8')
            self.csv_writer = csv.writer(self.csv_file)

            # 写入表头
            headers = [
                '时间', '变频器转速(RPM)', '设定转速(RPM)', '变频器电压(V)',
                '变频器电流(A)', '变频器功率(kW)', '变频器转矩(%)',
                '转矩仪转矩(Nm)', '转矩仪转速(RPM)', '转矩仪功率(W)'
            ]
            self.csv_writer.writerow(headers + self.config.modbus_head)

            self.log_message('info', f'开始记录数据到 {self.csv_filename}')
        except Exception as e:
            self.log_message('error', f'创建CSV文件失败: {str(e)}')
            self._close_csv_writer()

    # 关闭CSV写入器
    def _close_csv_writer(self):
        """关闭CSV写入器"""
        if self.csv_file:
            try:
                self.csv_file.close()
            except Exception as e:
                self.log_message('error', f'关闭CSV文件失败: {str(e)}')
            finally:
                self.csv_file = None
                self.csv_writer = None

    # 写入一行数据到CSV
    def _write_to_csv(self, record):
        """写入一行数据到CSV"""
        if self.csv_writer:
            try:
                self.csv_writer.writerow(record)
                self.csv_file.flush()  # 立即写入磁盘
            except Exception as e:
                self.log_message('error', f'写入CSV失败: {str(e)}')
                self._close_csv_writer()

    # # 保存采集的数据
    # def _save_collected_data(self):
    #     """保存采集的数据"""
    #     if hasattr(self, 'data_buffer') and self.data_buffer:
    #         try:
    #             timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    #             filename = f"motor_data_{timestamp}.xlsx"
    #
    #             wb = Workbook()
    #             ws = wb.active
    #             # 添加表头
    #             headers = [
    #                 '时间', '变频器转速(RPM)', '设定转速(RPM)', '变频器电压(V)',
    #                 '变频器电流(A)', '变频器功率(kW)', '变频器转矩(%)',
    #                 '转矩仪转矩', '转矩仪转速', '转矩仪功率', '运行状态'
    #             ]
    #             ws.append(headers)
    #
    #             # 添加数据行
    #             for row in self.data_buffer:
    #                 ws.append(row)
    #
    #             wb.save(filename)
    #             self.log_message('info', f'数据已保存到 {filename}')
    #         except Exception as e:
    #             self.log_message('error', f'保存数据失败: {str(e)}')
    #         finally:
    #             self.data_buffer = []  # 清空缓冲区

    # 更新数据显示
    def update_data_display(self, data):
        """更新数据显示"""
        # 更新变频器数据显示
        self.ui.ledoutrot.display(data['speed'])
        self.ui.ledoutvot.display(data['voltage'])
        self.ui.ledoutcur.display(data['current'])
        self.ui.ledoutpow.display(data['power'])
        self.ui.ledouttor.display(data['torque'])

        # 更新转矩仪数据显示
        self.ui.ledreadtor.display(data['torque_meter_torque'])
        self.ui.ledreadrot.display(data['torque_meter_speed'])
        self.ui.ledreadpwr.display(data['torque_meter_power'])

        # 更新转矩仪数据显示
        self.ui.ledCH0.display(data['ch0'])
        self.ui.ledCH1.display(data['ch1'])
        self.ui.ledCH2.display(data['ch2'])
        self.ui.ledCH3.display(data['ch3'])
        self.ui.ledCH4.display(data['ch4'])
        self.ui.ledCH5.display(data['ch5'])
        self.ui.ledCH6.display(data['ch6'])
        self.ui.ledCH7.display(data['ch7'])

        # 更新运行状态
        status = data['status']
        if 0 <= status < len(self.run_status_text):
            self.ui.labisrun.setText(self.run_status_text[status])

        # 如果需要保存数据
        if self.ui.cboxdaq.isChecked():
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
            record = [
                timestamp,
                data['speed'],
                data.get('set_speed', 0),  # 如果没有设定转速则使用0
                data['voltage'],
                data['current'],
                data['power'],
                data['torque'],
                data['torque_meter_torque'],
                data['torque_meter_speed'],
                data['torque_meter_power'],
                data['ch0'],
                data['ch1'],
                data['ch2'],
                data['ch3'],
                data['ch4'],
                data['ch5'],
                data['ch6'],
                data['ch7'],

            ]
            # self.data_buffer.append(record)
            self._write_to_csv(record)  # 直接写入CSV

    # 发送自定义命令
    def send_custom_command(self):
        """发送自定义命令"""
        command = self.ui.intgra.text().strip()
        if not command:
            self.log_message('error', '请输入命令')
            return

        if len(command) % 2 != 0:
            self.log_message('error', '命令长度必须为偶数')
            return

        self.log_message('info', f'发送命令: {command}')
        success, response = self.send_command(command)

        if success:
            self.log_message('info', f'收到响应: {response}')
        else:
            self.log_message('error', '命令发送失败')

    # 清空命令显示
    def clear_command_display(self):
        """清空命令显示"""
        self.ui.wrigra.clear()

    # 记录日志消息
    def log_message(self, level, message):
        """记录日志消息"""
        levels = {
            'info': Qt.black,
            'warning': Qt.darkYellow,
            'error': Qt.red,
            'debug': Qt.gray
        }

        color = levels.get(level.lower(), Qt.black)
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]

        self.ui.wrigra.setTextColor(color)
        self.ui.wrigra.append(f"[{timestamp}] {level.upper()}: {message}")
        self.ui.wrigra.setTextColor(Qt.black)

        # 可选：同时输出到控制台
        print(f"[{timestamp}] {level.upper()}: {message}")

    # 关闭窗口事件处理
    def closeEvent(self, event):
        """关闭窗口事件处理"""
        if self.thread and self.thread.isRunning():
            self.thread.stop()
            self.thread.wait(2000)

        self._close_socket()
        super().closeEvent(event)


# 数据采集线程
class DataCollectionThread(QThread):
    """数据采集线程"""
    data_ready = pyqtSignal(dict)

    def __init__(self, interval, controller):
        super().__init__()
        self.interval = interval
        self.controller = controller
        self._running = False
        self.request = bytes.fromhex("00 00 00 00 00 06 00 04 01 01 00 08")
        self.sock2 = None  # 第二socket,用于连接modbus采集卡
        self.config = MotorConfig()
        if self.config.usesocket2:
            self._connect_to_modbus()  # 物理量采集连接

    # 建立socket2连接
    def _connect_to_modbus(self):
        """建立socket连接"""
        try:
            self.sock2 = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.sock2.setsockopt(socket.IPPROTO_TCP, socket.TCP_NODELAY, 1)
            self.sock2.settimeout(0.5)
            self.sock2.connect((self.config.ip_address2, self.config.port2))

            MotorController.log_message(self.controller, 'info', '物理量采集连接成功')
            return True
        except Exception as e:
            MotorController.log_message(self.controller, 'error', f'物理量采集连接失败: {str(e)}')
            self._close_socket()
            return False

    # 安全关闭socket
    def _close_socket(self):
        """安全关闭socket"""
        if self.sock2:
            try:
                self.sock2.close()
            except:
                pass
            finally:
                self.sock2 = None

    def run(self):
        self._running = True
        next_time = time.time()

        while self._running:
            try:
                # 收集数据
                data = self._collect_data()
                if data:
                    self.data_ready.emit(data)

                # 精确的时间控制
                next_time += self.interval
                sleep_time = next_time - time.time()
                if sleep_time > 0:
                    time.sleep(sleep_time)
                else:
                    print("Warning: Data collection can't keep up with interval")
                    MotorController.log_message(self.controller, 'warning', f'采集阻塞')
            except Exception as e:
                print(f"Data collection error: {e}")
                MotorController.log_message(self.controller, 'error', f'采集失败: {e}')
                time.sleep(1)  # 出错时短暂等待

    # 收集电机数据
    def _collect_data(self):
        """收集电机数据"""
        data = {}

        # 读取电机参数
        success, response = self.controller.send_command('010370000007', '0103')

        if success:
            try:
                read_data = struct.unpack('>hhhhhhh', bytes.fromhex(response)[3:17])
                data['speed'] = round(read_data[0] * 0.6 * self.controller.motor_params['rotation_ratio'])
                data['voltage'] = read_data[3]
                data['current'] = read_data[4] / 100
                data['power'] = read_data[5] / 10
                data['torque'] = read_data[6] / 10
            except Exception as e:
                print(f"解析电机参数出错: {e}")

        # # 读取运行状态
        if data['speed'] >= 5:
            data['status'] = 1
        else:
            data['status'] = 0
        # 读取转矩仪数据
        success, response = self.controller.send_command('020300000006', '0203')

        if success:
            try:
                read_data = struct.unpack('>iii', bytes.fromhex(response)[3:15])
                data['torque_meter_torque'] = read_data[0] / 100
                data['torque_meter_speed'] = read_data[1] / 10
                data['torque_meter_power'] = read_data[2] / 100
            except Exception as e:
                print(f"解析转矩仪数据出错: {e}")


        for i in range(8):
            data[f'ch{i}'] = -1

        if self.config.usesocket2:
            self.sock2.sendall(self.request)
            # 接收响应
            response = self.sock2.recv(1024)
            if len(response) >= 25:
                read_data = struct.unpack('>8H', response[9:25])
                for i in range(8):
                    data[f'ch{i}'] = read_data[i] * (self.config.modbus_max[i] - self.config.modbus_min[i]) / 65536 + \
                                     self.config.modbus_min[i]
                    data[f'ch{i}']=round(data[f'ch{i}'], 6)
                print(data[f'ch1'])
                print(type(data[f'ch1']))
                # print([data[f'ch{i}'] for i in range(8) ])

        return data if data else None

    # 停止线程
    def stop(self):
        """停止线程"""
        # if self.config.usesocket2:
        #     self.sock2.close()
        self._running = False


class CRCHelper:
    """CRC校验工具类"""

    # 计算CRC16校验码
    @staticmethod
    def calculate_crc(data):
        """计算CRC16校验码"""
        crc = 0xFFFF
        for byte in data:
            crc ^= byte
            for _ in range(8):
                if crc & 0x0001:
                    crc >>= 1
                    crc ^= 0xA001
                else:
                    crc >>= 1
        return crc.to_bytes(2, 'little')

    # 验证CRC校验码
    @staticmethod
    def verify_crc(data):
        """验证CRC校验码"""
        if len(data) < 2:
            return False, None

        payload = data[:-2]
        received_crc = data[-2:]
        calculated_crc = CRCHelper.calculate_crc(payload)
        return received_crc == calculated_crc, payload

    # 添加CRC校验码到数据
    @staticmethod
    def add_crc(data):
        """添加CRC校验码到数据"""
        crc = CRCHelper.calculate_crc(data)
        return data + crc


# 电机配置类
class MotorConfig:
    """电机配置类"""

    def __init__(self):
        self.ip_address = "192.168.1.122"
        self.port = 8802
        self.ip_address2 = "192.168.1.220"
        self.port2 = 502
        self.max_speed = 3000
        self.rotation_ratio = 1
        self.sample_interval = 0.1
        self.usesocket2 = 1
        self.modbus_head = ["温度1【℉】", "压力1【bar】", "流量1【sccm】", "振动1【mm/s^2】",
                            "温度2【℃】", "压力2【mpa】", "流量2【slm】", "振动2【mm/s】"]
        self.modbus_min = [0, 0, 0, 0, 0, 0, 0, 0]
        self.modbus_max = [100, 100, 100, 100, 100, 100, 100, 100]
        self.load_from_file("config.json")

    def save_to_file(self, filename):
        """保存配置到文件"""
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(self.__dict__, f, ensure_ascii=False, indent=4)

    # def load_from_file(self, filename):
    #     """从文件加载配置"""
    #     with open(filename, 'r') as f:
    #         data = json.load(f)
    #         self.__dict__.update(data)

    def load_from_file(self, filename):
        """
        从文件加载配置
        如果文件不存在，则保存当前配置为默认值
        """
        try:
            with open(filename, 'r', encoding='utf-8') as f:
                data = json.load(f)
                self.__dict__.update(data)
        except FileNotFoundError:
            print(f"配置文件 {filename} 不存在，创建默认配置")
            self.save_to_file(filename)  # 保存当前配置
        except json.JSONDecodeError:
            print(f"配置文件 {filename} 格式错误，创建默认配置")
            self.save_to_file(filename)  # 保存当前配置


if __name__ == '__main__':
    app = QApplication(sys.argv)

    try:
        controller = MotorController()
        controller.ui.show()
        sys.exit(app.exec_())
    except Exception as e:
        print(f"程序启动失败: {str(e)}")
        QMessageBox.critical(None, "错误", f"程序启动失败: {str(e)}")
