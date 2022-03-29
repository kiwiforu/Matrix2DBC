# -*- coding: utf-8 -*

#====================import======================
import os
import xlrd
import collections
import math
#====================import======================


#====================defined=====================
DBC_FILE_COLUMN = ['message','id','dlc','cycletime','msgsendtype','format','singal','startbit','length[Bit]','ByteOrder','ValueType','InitialValue','Factor','Offset','Min','Max','Unit', 'Tx', 'Rx']

NEW_SYMBOLS =\
[
    'NS_DESC_', 'CM_', 'BA_DEF_', 'BA_', 'VAL_', 'CAT_DEF_', 'CAT_',
    'FILTER', 'BA_DEF_DEF_', 'EV_DATA_', 'ENVVAR_DATA_', 'SGTYPE_',
    'SGTYPE_VAL_', 'BA_DEF_SGTYPE_', 'BA_SGTYPE_', 'SIG_TYPE_REF_',
    'VAL_TABLE_', 'SIG_GROUP_', 'SIG_VALTYPE_', 'SIGTYPE_VALTYPE_',
    'BO_TX_BU_', 'BA_DEF_REL_', 'BA_REL_', 'BA_DEF_DEF_REL_',
    'BU_SG_REL_', 'BU_EV_REL_', 'BU_BO_REL_', 'SG_MUL_VAL_'
]
DEFAULT_BA_DEF =\
[
# object type      name                      value type     min max     default                 value range
    [ "BO_",  "VFrameFormat",               "Enumeration",  "", "", "StandardCAN",              ["StandardCAN","ExtendedCAN","reserved","reserved","reserved","reserved","reserved","reserved","reserved","reserved","reserved","reserved","reserved","reserved","StandardCAN_FD","ExtendedCAN_FD"]],
    [ "BO_",  "GenMsgStartDelayTime" ,      "Integer",      0,  0,      0,                      []],
    [ "BO_",  "GenMsgDelayTime" ,           "Integer",      0,  0,      0,                      []],
    [ "BO_",  "GenMsgNrOfRepetition" ,      "Integer",      0,  0,      0,                      []],
    [ "BO_",  "GenMsgCycleTimeFast" ,       "Integer",      0,  0,      0,                      []],
    [ "BO_",  "GenMsgCycleTime" ,           "Integer",      0,  0,      0,                      []],
    [ "BO_",  "GenMsgSendType" ,            "Enumeration",  "", "", "Cyclic",                   ["Cyclic","NotUsed","NotUsed","NotUsed","NotUsed","NotUsed","NotUsed","IfActive","NoMsgSendType","NotUsed"]],
    [ "SG_",  "GenSigStartValue" ,          "Integer",      0,  0,      0,                      []],
    [ "SG_",  "GenSigInactiveValue" ,       "Integer",      0,  0,      0,                      []],
    [ "SG_",  "GenSigCycleTimeActive" ,     "Integer",      0,  0,      0,                      []],
    [ "SG_",  "GenSigCycleTime" ,           "Integer",      0,  0,      0,                      []],
    [ "SG_",  "GenSigSendType" ,            "Enumeration",  "", "",     "Cyclic",               ["Cyclic","OnWrite","OnWriteWithRepetition","OnChange","OnChangeWithRepetition","IfActive","IfActiveWithRepetition","NoSigSendType","NotUsed","NotUsed","NotUsed","NotUsed","NotUsed"]],
    [ "",     "Baudrate" ,                  "Integer",      0,  1000000, 500000,                []],
    [ "",     "BusType" ,                   "String",       "", "",     "",                     []],
    [ "",     "NmType" ,                    "String",       "", "",     "",                     []],
    [ "",     "Manufacturer" ,              "String",       "", "",     "",                     []],
    [ "BO_",  "TpTxIndex" ,                 "Integer",      0,  255,    0,                      []],
    [ "BU_",  "NodeLayerModules" ,          "String",       "", "",     "CANoeILNLVector.dll",  []],
    [ "BU_",  "NmStationAddress",           "Hex",          0x0, 0x7FF, 0x400,                  []],
    [ "BU_",  "NmNode" ,                    "Enumeration",  "", "",     "no",                   ["no","yes"]],
    [ "BO_",  "NmMessage" ,                 "Enumeration",  "", "",     "no",                   ["no","yes"]],
    [ "",     "NmAsrWaitBusSleepTime" ,     "Integer",      0,  65535,  1500,                   []],
    [ "",     "NmAsrTimeoutTime" ,          "Integer",      1,  65535,  2000,                   []],
    [ "",     "NmAsrRepeatMessageTime" ,    "Integer",      0,  65535,  3200,                   []],
    [ "BU_",  "NmAsrNodeIdentifier",        "Hex",          0x0, 0xFF,  0x50,                   []],
    [ "BU_",  "NmAsrNode" ,                 "Enumeration",  "", "",     "no",                   ["no","yes" ]],
    [ "",     "NmAsrMessageCount" ,         "Integer",      1,  256,    128,                    []],
    [ "BO_",  "NmAsrMessage" ,              "Enumeration",  "", "",     "no",                   ["no","yes"]],
    [ "BU_",  "NmAsrCanMsgReducedTime" ,    "Integer",      1 , 65535,  320,                    []],
    [ "",     "NmAsrCanMsgCycleTime" ,      "Integer",      1,  65535,  640,                    []],
    [ "BU_",  "NmAsrCanMsgCycleOffset" ,    "Integer",      0,  65535,  0,                      []],
    [ "",     "NmAsrBaseAddress",           "Hex",          0x0, 0x7FF, 0x500,                  []],
    [ "BU_",  "ILUsed" ,                    "Enumeration",  "", "",     "no",                   ["no","yes"]],
    [ "",     "ILTxTimeout" ,               "Integer",      0,  65535,  0,                      []],
    [ "SG_",  "GenSigTimeoutValue" ,        "Integer",      0,  65535,   0,                      []],
    [ "SG_",  "GenSigTimeoutTime" ,         "Integer",      0,  65535,  0,                      []],
    [ "BO_",  "GenMsgILSupport" ,           "Enumeration",  "", "",     "yes",                  ["no","yes"]],
    [ "BO_",  "GenMsgFastOnStart" ,         "Integer",      0,  65535,  0,                      []],
    [ "BO_",  "DiagUudtResponse" ,          "Enumeration",  "", "",     "false",                ["false","true"]],
    [ "BO_",  "DiagUudResponse" ,           "Enumeration",  "", "",     "false",                ["false","True"]],
    [ "BO_",  "DiagState" ,                 "Enumeration",  "", "",     "no",                   ["no","yes"]],
    [ "BO_",  "DiagResponse" ,              "Enumeration",  "", "",     "no",                   ["no","yes"]],
    [ "BO_",  "DiagRequest" ,               "Enumeration",  "", "",     "no",                   ["no","yes"]],
    [ "",     "DBName" ,                    "String",       "", "",     "",                     []],
    [ "SG_",  "SystemSignalLongSymbol" ,    "String",       "", "",     "",                     []],
]

DEFAULT_PARAMENT =\
[
# objetc type  name                      value
    ['BA_', 'Manufacturer',             'Vector'],
    ['BA_', 'NmType',                   'NmAsr'],
    ['BA_', 'BusType',                  'CAN FD'],
    ['BA_', 'Baudrate',                 '500000'],
    ['BA_', 'NmAsrWaitBusSleepTime',    '2000'],
    ['BA_', 'DBName',                   ''],
]

MESSAGE_PARAMENT =\
[
    ['BA_', 'GenMsgSendType',   'BO_'],
    ['BA_', 'GenMsgCycleTime',  'BO_'],
    ['BA_', 'VFrameFormat',     'BO_'],
]

BTYE_ORDER_DEF = \
{
    'Motorola'  : '0',
    'Intel'     : '1'
}

VALUE_TYPE_DEF = \
{
    'Unsigned'  :'+',
    'Signed'    :'-'
}

SEND_TYPE_LIST = ["Cyclic","NotUsed","NotUsed","NotUsed","NotUsed","NotUsed","NotUsed","IfActive","NoMsgSendType","NotUsed"]
FORMAT_LIST = ["StandardCAN","ExtendedCAN","reserved","reserved","reserved","reserved","reserved","reserved","reserved","reserved","reserved","reserved","reserved","reserved","StandardCAN_FD","ExtendedCAN_FD"]

#====================defined=====================

#====================function====================

# ===================================================================
#     Method      : create_dbc_file
#
#     Description :
#           This method is create dbc file
#     Parameters  : fpath,
#                   name
#     Returns     : None
# ===================================================================
def create_dbc_file(fpath, name):
    fp = open(fpath + "\\" + name + ".dbc", 'w')
    fp.write(build_dbc_default_start())
    fp.write("\nBS_:\n")
    fp.close()

# ===================================================================
#     Method      : get_column_index
#
#     Description :
#           This method is get needed column title index
#     Parameters  : list_names,
#                   col_name
#     Returns     : index list
# ===================================================================
def get_column_index(list_names, col_name):
    index_list = []
    for list_name in list_names:
        if list_name in col_name:
            index_list.append(col_name.index(list_name))
    return index_list

# ===================================================================
#     Method      : read_matrix_file
#
#     Description :
#           This method is read matrix information from xlsx file
#     Parameters  : fpath
#     Returns     : matrix text
# ===================================================================
def read_matrix_file(fpath):
    dbc_data = []
    data = xlrd.open_workbook(fpath)
    sheet = data.sheets()[0]
    line_num = sheet.nrows
    column_titles = sheet.row_values(0)
    index_list = get_column_index(DBC_FILE_COLUMN, column_titles)
    for i in range(1, line_num):
        row = sheet.row_values(i)
        dbc_data.append(row)
    return dbc_data, index_list

# ===================================================================
#     Method      : build_message_info
#
#     Description :
#           This method is build message information
#     Parameters  : data
#     Returns     : message line
# ===================================================================
def build_dbc_default_start():
    text = 'VERSION \"\"\n\n\n' + 'NS_ :\n'
    for symbol in NEW_SYMBOLS:
        text += "	" + symbol + "\n"
    return text

# ===================================================================
#     Method      : find_keys
#
#     Description :
#           This method is find elements in list
#     Parameters  : list
#     Returns     : keys
# ===================================================================
def find_keys(list):
    keys = str(collections.Counter(list).keys())
    keys = keys.replace("dict_keys(", "", -1)
    keys = keys.replace(")", "", -1)
    keys = eval(keys)
    return keys

# ===================================================================
#     Method      : build_Ecu_list
#
#     Description :
#           This method is build Ecu list BU_: &
#     Parameters  : data
#     Returns     : text
# ===================================================================
def build_Ecu_list(data):
    Tx_list = []
    Rx_list = []
    for info in data:
        Tx_list.append(info[-2])
        Rx_list.append(info[-1])
    Tx = find_keys(Tx_list)
    Rx = find_keys(Rx_list)
    text = "BU_: "
    for ecu in Tx:
        text += ecu + " "
    for ecu in Rx:
        text += ecu + " "
    return text

# ===================================================================
#     Method      : build_BO_info
#
#     Description :
#           This method is build BO_ information
#     Parameters  : line
#     Returns     : text
# ===================================================================
def build_BO_info(line):
    line[0] = line[0].replace(" ", "", -1)
    if line[-2] == '':
        line[-2] = 'Vector__XXX'
    text = '\nBO_ ' + str(int(line[1], 16)) + " " + line[0] + ": " + str(int(line[2])) + " " + line[-2]
    return text

# ===================================================================
#     Method      : build_SG_info
#
#     Description :
#           This method is build SG_ information
#     Parameters  : line
#     Returns     : text
# ===================================================================
def build_SG_info(line):
    line[6] = line[6].replace(" ", "", -1)
    if line[14] == '':
        line[14] = 0
        line[14] = int(line[14])
    else:
        if line[14] % 1 == 0:
            line[14] = int(line[14])

    if line[15] == '':
        line[15] = math.pow(2, int(line[8])) - 1
        line[15] = int(line[15])
    else:
        if line[15] % 1 == 0:
            line[15] = int(line[15])

    if line[-1] == '':
        line[-1] = 'Vector__XXX'
    text = ' SG_ ' + line[6] + " : " + str(int(line[7])) + '|' + str(int(line[8])) + '@'+BTYE_ORDER_DEF[line[9]]+VALUE_TYPE_DEF[line[10]]+'('+str(int(line[12])) + ',' + str(int(line[13])) + ') [' + str(line[14]) + '|' + str(line[15]) + '] \"' + str(line[-3]) + '\"  ' + line[-1]
    return text

# ===================================================================
#     Method      : build_message_info
#
#     Description :
#           This method is build message information
#     Parameters  : data
#     Returns     : message line
# ===================================================================
def build_message_info(data, list, file):
    for info in data:
        if info[0] != "":
            file.write(build_BO_info(info) + '\n')
            print(build_BO_info(info))
            file.write(build_SG_info(info) + '\n')
            print(build_SG_info(info))
        else:
            file.write(build_SG_info(info) + '\n')
            print(build_SG_info(info))

# ===================================================================
#     Method      : build_default_define
#
#     Description :
#           This method is build default define
#     Parameters  : array
#                   file
#     Returns     : None
# ===================================================================
def build_default_define(array, file):
    for element in array:
        if element[2] == 'Enumeration':
            value_type = 'ENUM'
            element[6] = str(element[6]).replace('[', '', -1)
            element[6] = element[6].replace(']', '', -1)
            element[6] = element[6].replace('\'', '\"', -1)
            text = 'BA_DEF_ ' + element[0] + '  \"' + element[1] + '\" ' + value_type + '  ' + element[6]+';\n'
        elif element[2] == 'Integer':
            value_type = 'INT'
            text = 'BA_DEF_ ' + element[0] + '  \"' + element[1] + '\" ' + value_type + '  ' + str(element[3])+' '+str(element[4])+';\n'
        elif element[2] == 'String':
            value_type = 'STRING'
            text = 'BA_DEF_ ' + element[0] + '  \"' + element[1] + '\" ' + value_type+' ;\n'
        elif element[2] == 'Hex':
            value_type = 'HEX'
            text = 'BA_DEF_ ' + element[0] + '  \"' + element[1] + '\" ' + value_type + '  ' + str(element[3])+' '+str(element[4])+';\n'
        fp.write(text)
        print(text)
    for element in array:
        text = 'BA_DEF_DEF_  \"'+element[1]+'\" '
        if element[2] == 'Enumeration' or element[2] == 'String':
            text += '\"'+element[5]+'\";\n'
        elif element[2] == 'Integer' or element[2] == 'Hex':
            text += str(element[5])+';\n'
        fp.write(text)
        print(text)

# ===================================================================
#     Method      : build_default_parament
#
#     Description :
#           This method is build default parament
#     Parameters  : array
#                   name
#                   file
#     Returns     : None
# ===================================================================
def build_default_parament(array, name, file):
    for element in array:
        if True == element[2].isdigit():
            text = element[0]+' \"'+element[1]+'\" '+element[2]+';\n'
        elif '' == element[2]:
            text = element[0]+' \"'+element[1]+'\" \"'+name+'\";\n'
        else:
            text = element[0]+' \"'+element[1]+'\" \"'+element[2]+'\";\n'
        fp.write(text)
        print(text)

# ===================================================================
#     Method      : find_signal_parament_index
#
#     Description :
#           This method is build signal parament
#     Parameters  : array
#                   str1
#                   str2
#     Returns     : idx
# ===================================================================
def find_signal_parament_index(array, str1, str2):
    for element in array:
        if str1 in element:
            idx = element[6].index(str2)
            print(idx)
    return idx

# ===================================================================
#     Method      : find_messgae_idx
#
#     Description :
#           This method is build which line the message belong
#     Parameters  : data
#     Returns     : idx_list
#                   canid_list
# ===================================================================
def find_message_idx(data):
    idx_list = []
    canid_list = []
    cycle_time_list = []
    send_type_list = []
    format_list = []
    for info in data:
        if info[0] != '':
            idx_list.append(data.index(info))
            canid_list.append(info[1])
            cycle_time_list.append(info[3])
            send_type_list.append(info[4])
            format_list.append(info[5])

    return idx_list, canid_list, cycle_time_list, send_type_list, format_list
# ===================================================================
#     Method      : build_signal_parament
#
#     Description :
#           This method is build signal parament
#     Parameters  : data
#                   file
#     Returns     : None
# ===================================================================
def build_signal_parament(data, array, file):
    idx_list, canid_list, cycle_time_list, send_type_list, format_list = find_message_idx(data)
    cnt = 0
    for idx in idx_list:
        id = str(int(canid_list[cnt], 16))
        for element in array:
            text = element[0]+' \"'+element[1]+'\" '+element[2]+' '
            if 'GenMsgSendType' == element[1]:
                text += id+' '+str(SEND_TYPE_LIST.index(send_type_list[cnt]))+';\n'
            elif 'GenMsgCycleTime' == element[1]:
                text += id+' '+str(int(cycle_time_list[cnt]))+';\n'
            elif 'VFrameFormat' == element[1]:
                text += id+' '+str(FORMAT_LIST.index(format_list[cnt]))+';\n'
            print(text)
            fp.write(text)
        cnt +=1


#====================function====================

#====================main========================
if __name__ == '__main__':
    path = input("input xlsx file path:")
    name = input("please enter the name:")
    data, index_list = read_matrix_file(path + '\dbc.xlsx')

    fp = open(path + '\\' + name + '.dbc', 'w')
    fp.write(build_dbc_default_start())
    fp.write("\nBS_:\n\n")
    fp.write(build_Ecu_list(data) + '\n')
    build_message_info(data, index_list, fp)
    fp.write('\n')
    build_default_define(DEFAULT_BA_DEF, fp)
    build_default_parament(DEFAULT_PARAMENT, name, fp)
    build_signal_parament(data, MESSAGE_PARAMENT, fp)
    fp.close()



    #create_dbc_file(r'C:\Users\kiwi\Desktop\data', 'T5')
#====================main========================