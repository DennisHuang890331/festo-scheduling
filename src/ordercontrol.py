import sys
sys.path.insert(0, "..")
from opcua import Client, ua
import socket
import time

            
def TwoOrders(orders):
    firstCarLeavingThePressure = False
    while True:
        if P_OPos_1.get_value() == 1 and P_bg27.get_value() == True:
            firstCarLeavingThePressure = True
            print("firstCarLeavingThePressure",firstCarLeavingThePressure)
            
        if P_OPos_1.get_value() == 2 and firstCarLeavingThePressure == True  :
            P_automatic.set_attribute(ua.AttributeIds.Value, ua.DataValue(False))
            P_manual.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))
            #加壓輸送帶保持運輸
            P_belt.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))
            firstCarLeavingThePressure = False
            print("firstCarLeavingThePressure",firstCarLeavingThePressure)
        elif M_oPos_2.get_value() == 1:
            P_manual.set_attribute(ua.AttributeIds.Value, ua.DataValue(False))
            P_reset.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))
            time.sleep(1)
            P_automatic.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))
            exit()

def ThreeOrders(orders):
     #布林值們
    lastCarGotoCheck = False
    firstCarGotoHeat  = False
    LastCarGotoHeat  = False
    firstCarTakingAGV = False
    firstCarLeavingThePressure = False
    takingAGVtoBack = False
    takingAGVtoGO = False
    while True:
        """WORK HERE"""
            #"""WORK HERE"""
        #僅適用4~9單
        if LastCarGotoHeat == True:
            #倉儲auto
            S_manual.set_attribute(ua.AttributeIds.Value, ua.DataValue(False))
            S_reset.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))
            time.sleep(10)
            S_automatic.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))
            print('The end')
            exit()
        elif firstCarGotoHeat == True:
            #讀取下蓋RFID 若為最後一單則將倉儲右側氣壓閥升起 擋住第一單
            if b_OPos_1.get_value() == orders:
                S_mb20.set_attribute(ua.AttributeIds.Value, ua.DataValue(False))
                print("倉儲右側已擋起")
            #讀取加壓RFID 若為最後一單 且 倉儲至視覺輸送帶被觸發時 代表最後一單準備去加熱 可將倉儲轉回auto
            if V_oPos_2.get_value() == orders and P_OPos_1.get_value() == orders:
                LastCarGotoHeat = True
                print("LastCarGotoHeat",LastCarGotoHeat)
            else:
                print("LastCarGotoHeat",LastCarGotoHeat)
            #放不放車過                
            #量測左車道BG偵測到上車的False才讓下一台車進入公車站
            if V_oPos_2.get_value() == 1:
                if V_bg42.get_value() == True:
                    firstCarTakingAGV = True
                if firstCarTakingAGV == True:
                    if V_bg42.get_value() == False:
                        firstCarTakingAGV = False
                        P_mb20.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))
                elif P_bg21.get_value() == False:
                    P_mb20.set_attribute(ua.AttributeIds.Value, ua.DataValue(False))

            else:
                if M_bg42.get_value() == True:
                    takingAGVtoBack = True
                if takingAGVtoBack == True:
                    if M_bg42.get_value() == False:
                        takingAGVtoBack = False
                        S_mb30.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))
                elif S_bg31.get_value() == False:
                    S_mb30.set_attribute(ua.AttributeIds.Value, ua.DataValue(False))
        else:
            if lastCarGotoCheck == True:         
                if S_oPos_2.get_value() == 1:
                    S_automatic.set_attribute(ua.AttributeIds.Value, ua.DataValue(False))
                    S_manual.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))
                    #使右邊暢通
                    S_mb20.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))        
                    S_belt_1.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))
                        #使左邊輸送帶動
                    S_belt_2.set_attribute(ua.AttributeIds.Value, ua.DataValue(True)) 
                    firstCarGotoHeat = True
                    print("firstCarGotoHeat",firstCarGotoHeat)
            else:
                #當地一台車通過加壓，且離開時踩到輸送帶上的BG時，打開"離開加壓"旗標
                if P_OPos_1.get_value() == 1 and P_bg27.get_value() == True:
                    firstCarLeavingThePressure = True
                    print("firstCarLeavingThePressure",firstCarLeavingThePressure)
                    
                if P_OPos_1.get_value() == 2 and firstCarLeavingThePressure == True  :
                    P_automatic.set_attribute(ua.AttributeIds.Value, ua.DataValue(False))
                    P_manual.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))
                    #加壓輸送帶保持運輸
                    P_belt.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))
                    firstCarLeavingThePressure = False
                    print("firstCarLeavingThePressure",firstCarLeavingThePressure)
                #當第三台通過加熱後RFID 且BG轉為True時 將加壓切回自動
                elif M_oPos_2.get_value() == orders-2: #and M_bg21 == True:
                    P_manual.set_attribute(ua.AttributeIds.Value, ua.DataValue(False))
                    P_reset.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))
                    time.sleep(1)
                    P_automatic.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))
                    lastCarGotoCheck = True
                    print("lastCarGotoCheck",lastCarGotoCheck)
        #放不放車---------------------------------------------------------------------------
            #Solution 視覺RFIT等於1 且 候車BG等於true時，將等車旗標打開
            #旗標打開後，偵測到候車BG轉為Flase的時候，代表車子離開，則降氣壓筏讓下一台過來，旗標關閉。
            if V_oPos_2.get_value() == 1:
                if V_bg42.get_value() == True:
                    firstCarTakingAGV = True
                if firstCarTakingAGV == True:
                    if V_bg42.get_value() == False:
                        firstCarTakingAGV = False
                        P_mb20.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))
                elif P_bg21.get_value() == False:
                    P_mb20.set_attribute(ua.AttributeIds.Value, ua.DataValue(False))
            else:                                      #量測左車道偵測到上車才讓下一台車進入公車站
                # if M_bg21.get_value() == True:
                #     P_mb20.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))
                # else:
                #     P_mb20.set_attribute(ua.AttributeIds.Value, ua.DataValue(False))
                if M_bg42.get_value() == True:
                    takingAGVtoGO = True
                if takingAGVtoGO == True:
                    if M_bg42.get_value() == False:
                        takingAGVtoGO = False
                        P_mb20.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))
                elif P_bg21.get_value() == False:
                    P_mb20.set_attribute(ua.AttributeIds.Value, ua.DataValue(False))
            
def FourOrders(orders):
     #布林值們
    lastCarGotoCheck = False
    firstCarGotoHeat  = False
    LastCarGotoHeat  = False
    firstCarTakingAGV = False
    firstCarLeavingThePressure = False
    takingAGVtoBack = False
    takingAGVtoGO = False
    while True:
        """WORK HERE"""
            #"""WORK HERE"""
        #僅適用4~9單
        if LastCarGotoHeat == True:
            #倉儲auto
            S_manual.set_attribute(ua.AttributeIds.Value, ua.DataValue(False))
            S_reset.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))
            time.sleep(10)
            S_automatic.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))
            print('The end')
            exit()
        elif firstCarGotoHeat == True:
            #讀取下蓋RFID 若為最後一單則將倉儲右側氣壓閥升起 擋住第一單
            if b_OPos_1.get_value() == orders:
                S_mb20.set_attribute(ua.AttributeIds.Value, ua.DataValue(False))
                print("倉儲右側已擋起")
            #讀取加壓RFID 若為最後一單 且 倉儲至視覺輸送帶被觸發時 代表最後一單準備去加熱 可將倉儲轉回auto
            if V_oPos_2.get_value() == orders and P_OPos_1.get_value() == orders:
                LastCarGotoHeat = True
                print("LastCarGotoHeat",LastCarGotoHeat)
            else:
                print("LastCarGotoHeat",LastCarGotoHeat)
            #放不放車過                
            #量測左車道BG偵測到上車的False才讓下一台車進入公車站
            if M_bg42.get_value() == True:
                takingAGVtoBack = True
            if takingAGVtoBack == True:
                if M_bg42.get_value() == False:
                    takingAGVtoBack = False
                    S_mb30.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))
            elif S_bg31.get_value() == False:
                S_mb30.set_attribute(ua.AttributeIds.Value, ua.DataValue(False))
        else:
            if lastCarGotoCheck == True:         
                if S_oPos_2.get_value() == 2:
                    S_automatic.set_attribute(ua.AttributeIds.Value, ua.DataValue(False))
                    S_manual.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))
                    #使右邊暢通
                    S_mb20.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))        
                    S_belt_1.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))
                        #使左邊輸送帶動
                    S_belt_2.set_attribute(ua.AttributeIds.Value, ua.DataValue(True)) 
                    firstCarGotoHeat = True
                    print("firstCarGotoHeat",firstCarGotoHeat)
            else:
                #當地一台車通過加壓，且離開時踩到輸送帶上的BG時，打開"離開加壓"旗標
                if P_OPos_1.get_value() == 1 and P_bg27.get_value() == True:
                    firstCarLeavingThePressure = True
                    print("firstCarLeavingThePressure",firstCarLeavingThePressure)
                    
                if P_OPos_1.get_value() == 2 and firstCarLeavingThePressure == True  :
                    P_automatic.set_attribute(ua.AttributeIds.Value, ua.DataValue(False))
                    P_manual.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))
                    #加壓輸送帶保持運輸
                    P_belt.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))
                    firstCarLeavingThePressure = False
                    print("firstCarLeavingThePressure",firstCarLeavingThePressure)
                #當第三台通過加熱後RFID 且BG轉為True時 將加壓切回自動
                elif M_oPos_2.get_value() == orders-2: #and M_bg21 == True:
                    P_manual.set_attribute(ua.AttributeIds.Value, ua.DataValue(False))
                    P_reset.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))
                    time.sleep(1)
                    P_automatic.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))
                    lastCarGotoCheck = True
                    print("lastCarGotoCheck",lastCarGotoCheck)
        #放不放車---------------------------------------------------------------------------
            #Solution 視覺RFIT等於1 且 候車BG等於true時，將等車旗標打開
            #旗標打開後，偵測到候車BG轉為Flase的時候，代表車子離開，則降氣壓筏讓下一台過來，旗標關閉。
            if V_oPos_2.get_value() == 1:
                if V_bg42.get_value() == True:
                    firstCarTakingAGV = True
                if firstCarTakingAGV == True:
                    if V_bg42.get_value() == False:
                        firstCarTakingAGV = False
                        P_mb20.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))
                elif P_bg21.get_value() == False:
                    P_mb20.set_attribute(ua.AttributeIds.Value, ua.DataValue(False))
            else:                                      #量測左車道偵測到上車才讓下一台車進入公車站
                # if M_bg21.get_value() == True:
                #     P_mb20.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))
                # else:
                #     P_mb20.set_attribute(ua.AttributeIds.Value, ua.DataValue(False))
                if M_bg42.get_value() == True:
                    takingAGVtoGO = True
                if takingAGVtoGO == True:
                    if M_bg42.get_value() == False:
                        takingAGVtoGO = False
                        P_mb20.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))
                elif P_bg21.get_value() == False:
                    P_mb20.set_attribute(ua.AttributeIds.Value, ua.DataValue(False))

def FiveToNineOrders(orders):
    #OPCUA連線

    #布林值們
    lastCarGotoCheck = False
    firstCarGotoHeat  = False
    LastCarGotoHeat  = False
    firstCarTakingAGV = False
    firstCarLeavingThePressure = False
    takingAGVtoBack = False
    takingAGVtoGO = False
    while True:
        """WORK HERE"""
            #"""WORK HERE"""
        #僅適用4~9單
        if LastCarGotoHeat == True:
            #倉儲auto
            S_manual.set_attribute(ua.AttributeIds.Value, ua.DataValue(False))
            S_reset.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))
            time.sleep(10)
            S_automatic.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))
            print('The end')
            exit()
        elif firstCarGotoHeat == True:
            #讀取下蓋RFID 若為最後一單則將倉儲右側氣壓閥升起 擋住第一單
            if b_OPos_1.get_value() == orders:
                S_mb20.set_attribute(ua.AttributeIds.Value, ua.DataValue(False))
                print("倉儲右側已擋起")
            #讀取加壓RFID 若為最後一單 且 倉儲至視覺輸送帶被觸發時 代表最後一單準備去加熱 可將倉儲轉回auto
            if V_oPos_2.get_value() == orders and P_OPos_1.get_value() == orders:
                LastCarGotoHeat = True
                print("LastCarGotoHeat",LastCarGotoHeat)
            else:
                print("LastCarGotoHeat",LastCarGotoHeat)
            #放不放車過                
            #量測左車道BG偵測到上車的False才讓下一台車進入公車站
            if M_bg42.get_value() == True:
                takingAGVtoBack = True
            if takingAGVtoBack == True:
                if M_bg42.get_value() == False:
                    takingAGVtoBack = False
                    S_mb30.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))
            elif S_bg31.get_value() == False:
                S_mb30.set_attribute(ua.AttributeIds.Value, ua.DataValue(False))
        else:
            if lastCarGotoCheck == True:         
                if S_oPos_2.get_value() == 1:
                    S_automatic.set_attribute(ua.AttributeIds.Value, ua.DataValue(False))
                    S_manual.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))
                    #使右邊暢通
                    S_mb20.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))        
                    S_belt_1.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))
                        #使左邊輸送帶動
                    S_belt_2.set_attribute(ua.AttributeIds.Value, ua.DataValue(True)) 
                    firstCarGotoHeat = True
                    print("firstCarGotoHeat",firstCarGotoHeat)
            else:
                #當地一台車通過加壓，且離開時踩到輸送帶上的BG時，打開"離開加壓"旗標
                if P_OPos_1.get_value() == 1 and P_bg27.get_value() == True:
                    firstCarLeavingThePressure = True
                    print("firstCarLeavingThePressure",firstCarLeavingThePressure)
                    
                if P_OPos_1.get_value() == 2 and firstCarLeavingThePressure == True  :
                    P_automatic.set_attribute(ua.AttributeIds.Value, ua.DataValue(False))
                    P_manual.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))
                    #加壓輸送帶保持運輸
                    P_belt.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))
                    firstCarLeavingThePressure = False
                    print("firstCarLeavingThePressure",firstCarLeavingThePressure)
                #當第三台通過加熱後RFID 且BG轉為True時 將加壓切回自動
                elif M_oPos_2.get_value() == orders-2: #and M_bg21 == True:
                    P_manual.set_attribute(ua.AttributeIds.Value, ua.DataValue(False))
                    P_reset.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))
                    time.sleep(1)
                    P_automatic.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))
                    lastCarGotoCheck = True
                    print("lastCarGotoCheck",lastCarGotoCheck)
        #放不放車---------------------------------------------------------------------------
            #Solution 視覺RFIT等於1 且 候車BG等於true時，將等車旗標打開
            #旗標打開後，偵測到候車BG轉為Flase的時候，代表車子離開，則降氣壓筏讓下一台過來，旗標關閉。
            if V_oPos_2.get_value() == 1:
                if V_bg42.get_value() == True:
                    firstCarTakingAGV = True
                if firstCarTakingAGV == True:
                    if V_bg42.get_value() == False:
                        firstCarTakingAGV = False
                        P_mb20.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))
                elif P_bg21.get_value() == False:
                    P_mb20.set_attribute(ua.AttributeIds.Value, ua.DataValue(False))
            else:                                      #量測左車道偵測到上車才讓下一台車進入公車站
                # if M_bg21.get_value() == True:
                #     P_mb20.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))
                # else:
                #     P_mb20.set_attribute(ua.AttributeIds.Value, ua.DataValue(False))
                if M_bg42.get_value() == True:
                    takingAGVtoGO = True
                if takingAGVtoGO == True:
                    if M_bg42.get_value() == False:
                        takingAGVtoGO = False
                        P_mb20.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))
                elif P_bg21.get_value() == False:
                    P_mb20.set_attribute(ua.AttributeIds.Value, ua.DataValue(False))


    
Pressure = Client("opc.tcp://172.21.2.1:4840/")
Storage = Client("opc.tcp://172.21.3.1:4840/")
Visual = Client("opc.tcp://172.21.4.1:4840/")
Measure = Client("opc.tcp://172.21.6.1:4840/")
back = Client("opc.tcp://172.21.1.1:4840/")
try:
    Measure.connect()
    Visual.connect()
    Storage.connect()
    Pressure.connect()
    back.connect()

    #量測節點
    M_bg21 = Measure.get_node("ns=3;s=\"xW1_BG21\"")
    M_bg42 = Measure.get_node("ns=3;s=\"xW1_BG42\"")

    #2:準備進AGV
    M_oPos_2 = Measure.get_node("ns=3;s=\"dbRfidData\".\"ID1\".\"Mes\".\"iOPos\"")

    #視覺檢測節點
    V_bg42 = Visual.get_node("ns=3;s=\"xW1_BG42\"")
    V_oPos_2 = Visual.get_node("ns=3;s=\"dbRfidData\".\"ID1\".\"Mes\".\"iOPos\"")


    #倉儲點位 
    S_oPos_2 = Storage.get_node("ns=3;s=\"dbRfidData\".\"ID2\".\"Mes\".\"iOPos\"")      #belt_2
    S_belt_1 = Storage.get_node("ns=3;s=\"xQA1_A1\"")
    S_belt_2 = Storage.get_node("ns=3;s=\"xQA2_A1\"")
    S_automatic = Storage.get_node("ns=3;s=\"dbVar\".\"OpMode\".\"Auto\".\"xAct\"")
    S_manual = Storage.get_node("ns=3;s=\"dbVar\".\"OpMode\".\"Man\".\"xAct\"")
    S_mb30 = Storage.get_node("ns=3;s=\"xK1_MB30\"")        #Belt_2
    S_mb20 = Storage.get_node("ns=3;s=\"xK1_MB20\"")        #Belt_1
    S_bg31 = Storage.get_node("ns=3;s=\"xG1_BG31\"")
    S_reset = Storage.get_node("ns=3;s=\"dbVar\".\"OpMode\".\"Reset\".\"xAct\"")


    #加壓點位
    P_OPos_1 = Pressure.get_node("ns=3;s=\"dbRfidData\".\"ID1\".\"Mes\".\"iOPos\"")
    P_mb20 = Pressure.get_node("ns=3;s=\"xG1_MB20\"")
    P_belt = Pressure.get_node("ns=3;s=\"xQA1_A1\"")
    P_automatic = Pressure.get_node("ns=3;s=\"dbVar\".\"OpMode\".\"Auto\".\"xAct\"")
    P_manual = Pressure.get_node("ns=3;s=\"dbVar\".\"OpMode\".\"Man\".\"xAct\"")
    P_bg21 = Pressure.get_node("ns=3;s=\"xG1_BG21\"")
    P_bg27 = Pressure.get_node("ns=3;s=\"xG1_BG27\"")
    P_reset = Pressure.get_node("ns=3;s=\"dbVar\".\"OpMode\".\"Reset\".\"xAct\"")

    #下蓋點位
    b_OPos_1 = back.get_node("ns=3;s=\"dbRfidData\".\"ID1\".\"Mes\".\"iOPos\"")

    #主程式開始
    orders = int(input('單數:'))
    if orders>=5 and orders <= 9:
        FiveToNineOrders(orders)
    elif orders == 2:
        TwoOrders(orders)
    elif orders == 3:
        ThreeOrders(orders)
    elif orders == 4:
        FourOrders(orders)
    elif orders == 10:
        print('還沒寫好啦幹')
    else:
        print('去你的沒那麼多PCB啦幹')

except socket.error:
        print("連線錯誤，請確認是否已連上Taiwan-CP-Factory後再啟動程式")
finally:
    Measure.disconnect()
    Visual.disconnect()
    Pressure.disconnect()
    Storage.disconnect()
    back.disconnect()