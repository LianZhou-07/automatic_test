# automatic_test


```
硬件:
1.  用于Vin的直流电源 (IT6833)
2.  用于VCC的直流电源 (Rigol DP832)
3.  可编程负载 (Chroma 6310A)
4.  数字万用表 (Agilent 34970A)
5.  安装了NI-VISA的电脑。
```
```
设置:
-   将所有仪器连接到电脑。
-   确保在脚本中指定了正确的VISA资源地址。
-   DAQ（数据采集单元）分别在通道101, 102, 103, 和 104上测量
    Vin, V_sense_iin, Vout, 和 V_sense_iout。
-   输入电流(iin)是根据一个采样电阻上的电压(V_sense_iin)计算得出的。
-   输出电流(iout)是根据一个采样电阻上的电压(V_sense_iout)计算得出的。
```
