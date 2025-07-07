"""
批量扫测电源, 计算模块效率
-----------------------------------------------------------------------
  1. 通过 PyVISA 同时控制直流电源、辅助电源、电子负载和万用表
  2. 依次拉载多档电流，采集 Vin/Iin 和 Vout/Iout
  3. 计算系统效率与功率损耗并即时写入 Excel
  4. 支持自动生成带时间戳的文件名

• 软硬件依赖
  - Python（3.8以上）
  - pyvisa, pyvisa‑py, openpyxl
  - VISA 后端驱动 (NI/Keysight/Tek 等)
-----------------------------------------------------------------------
可根据实际仪表 SCPI 指令差异，通过修改 ADDRESS 与 CMD_* 常量自定义适配。
"""

from __future__ import annotations

import os
import time
from datetime import datetime
from pathlib import Path
from typing import Dict, Tuple

import openpyxl
import pyvisa

# 仪表地址（根据实际情况修改）
ADDRESS: Dict[str, str] = {
    "LOAD": "ASRL7::INSTR",                           # 负载串口
    "VIN_PS": "USB0::0x1698::0x0837::002000002608::INSTR",  # 直流电源
    "VCC_PS": "USB0::0x1AB1::0x0E11::DP8C193003485::INSTR", # 辅助电源
    "DAQ": "USB0::0x0957::0x2007::MY49029470::INSTR",       # Keysight DAQ（可按需增加）
}

# 分流电阻（Ω）——根据硬件实际值填写
SHUNT_IIN = 0.01  # @102 通道
SHUNT_IOUT = 0.001  # @104 通道

# 输出路径
OUTPUT_DIR = Path(r"E:/AutomaticEfficiencyForPlusCurrent")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# 辅助函数
def timestamp() -> str:
    """Return current datetime as yyyymmdd_HHMMSS string."""
    return datetime.now().strftime("%Y%m%d_%H%M%S")

def save_workbook(wb: openpyxl.Workbook, file_path: Path) -> None:
    """Save workbook and echo path."""
    wb.save(file_path)
    print(f"[INFO] 数据已保存到 {file_path}\n")


# 仪表相关封装
class Instruments:
    #封装所有仪表的打开/关闭与常用操作
    def __init__(self, rm: pyvisa.ResourceManager) -> None:
        self.rm = rm
        print("[INFO] VISA backend =>", rm)
        self.load = self._open("LOAD")
        self.vin_ps = self._open("VIN_PS")
        self.vcc_ps = self._open("VCC_PS")
        self.daq = self._open("DAQ")

        self._configure_load()

    # 私有工具
    def _open(self, key: str):
        try:
            inst = self.rm.open_resource(ADDRESS[key])
            print(f"[OK] Connected to {key}: {inst.query('*IDN?').strip()}")
            return inst
        except pyvisa.VisaIOError as err:
            raise RuntimeError(f"打开 {key} 失败: {err}") from err

    def _configure_load(self) -> None:
        """远程模式 + 选通道 3 + 恒流 0 A，保持关闭状态"""
        self.load.write("CONF:REM ON")
        self.load.write("CHAN 3")
        self.load.write("MODE CCH")
        self.load.write("CURR:STAT:L1 0")
        self.load.write("LOAD OFF")

    电源配置
    def setup_vin(self, voltage: float, current: float) -> None:
        self.vin_ps.write(f"SOUR:VOLT {voltage}")
        self.vin_ps.write(f"SOUR:CURR {current}")
        self.vin_ps.write("CONF:OUTP ON")

    def setup_vcc(self, voltage: float, current: float) -> None:
        self.vcc_ps.write("INST CH1")
        self.vcc_ps.write(f"VOLT {voltage}")
        self.vcc_ps.write(f"CURR {current}")
        self.vcc_ps.write("OUTP CH1,ON")

    # 负载控制
    def set_load_current(self, amp: float) -> None:
        self.load.write(f"CURR:STAT:L1 {amp}")

    def load_on(self):
        self.load.write("LOAD ON")

    def load_off(self):
        self.load.write("LOAD OFF")

    # 测量
    def measure(self) -> Tuple[float, float, float, float]:
        """Return Vin, Iin, Vout, Iout (float)."""
        vin = float(self.daq.query("MEAS:VOLT:DC? 100, 1E-4, (@101)"))
        iin_sense = float(self.daq.query("MEAS:VOLT? 2, 1E-5,(@102)"))
        vout = float(self.daq.query("MEAS:VOLT? 100, 1E-4,(@103)"))
        iout_sense = float(self.daq.query("MEAS:VOLT? 1, 1E-5,(@104)"))
        iin = iin_sense / SHUNT_IIN
        iout = iout_sense / SHUNT_IOUT
        return vin, iin, vout, iout

    # 关闭所有仪表
    def close_all(self):
        for inst in (self.load, self.vin_ps, self.vcc_ps, self.daq):
            try:
                inst.close()
            except Exception:  # noqa: BLE001
                pass
        print("所有仪表会话已关闭")


# 主流程

def main() -> None:
    # 交互输入
    vin_voltage = float(input("Vin 电压 (V): "))
    vin_current = float(input("Vin 限流 (A): "))

    vcc_voltage = float(input("Vcc 电压 (V): "))
    vcc_current = float(input("Vcc 限流 (A): "))

    iout_min = float(input("输出电流最小 (A): "))
    iout_max = float(input("输出电流最大 (A): "))
    iout_step = float(input("输出电流步进 (A): "))

    step_time = float(input("拉载稳定时间 (s): "))
    interval_time = float(input("间隔恢复时间 (s): "))

    # 初始化 Excel
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Efficiency_data"
    ws.append(["Vin(V)", "Iin(A)", "Vout(V)", "Iout(A)", "Eff_sys(%)", "Ploss_sys(W)"])

    file_path = OUTPUT_DIR / f"efficiency_{timestamp()}.xlsx"

    # VISA Resource Manager
    rm = pyvisa.ResourceManager()
    inst = Instruments(rm)

    try:
        # 配置电源
        inst.setup_vin(vin_voltage, vin_current)
        inst.setup_vcc(vcc_voltage, vcc_current)

        # 扫测循环
        current = iout_min
        while current <= iout_max + 1e-6:
            print(f"[TEST] Iout = {current:.3f} A")
            inst.set_load_current(current)
            inst.load_on()
            time.sleep(step_time)

            vin, iin, vout, iout = inst.measure()
            pin_sys = vin * iin
            pout = vout * iout
            ploss_sys = pin_sys - pout
            eff_sys = (pout / pin_sys) * 100 if pin_sys else 0

            # 写一行数据并立即保存
            ws.append([vin, iin, vout, iout, eff_sys, ploss_sys])
            save_workbook(wb, file_path)

            # 恢复
            inst.set_load_current(0)
            inst.load_off()
            time.sleep(interval_time)
            current += iout_step

    finally:
        # 关机收尾
        inst.close_all()
        wb.close()


if __name__ == "__main__":
    main()
