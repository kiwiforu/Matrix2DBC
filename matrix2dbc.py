# -*- coding: utf-8 -*

#====================import======================
import os
import xlrd
#====================import======================


#====================defined=====================
DBC_FILE_COLUMN = ['message','id','dlc','cycletime','msgsendtype','format','singal','startbit','length[Bit]','ByteOrder','ValueType','InitialValue','Factor','Offset','Min','Max','Unit']

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
    [ "BO_",  "GenMsgStartDelayTime",       "Integer",      0,  0,      0,                      []],
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
    [ "BU_",  "NmStationAddress"            "Hex",          0x0, 0x7FF, 0x400,                  []],
    [ "BU_",  "NmNode" ,                    "Enumeration",  "", "",     "no",                   ["no","yes"]],
    [ "BO_",  "NmMessage" ,                 "Enumeration",  "", "",     "no",                   ["no","yes"]],
    [ "",     "NmAsrWaitBusSleepTime" ,     "Integer",      0,  65535,  1500,                   []],
    [ "",     "NmAsrTimeoutTime" ,          "Integer",      1,  65535,  2000,                   []],
    [ "",     "NmAsrRepeatMessageTime" ,    "Integer",      0,  65535,  3200,                   []],
    [ "BU_",  "NmAsrNodeIdentifier",        "Hex",          0,  0xFF,   0x50,                   []],
    [ "BU_",  "NmAsrNode" ,                 "Enumeration",  "", "",     "no",                   ["no","yes" ]],
    [ "",     "NmAsrMessageCount" ,         "Integer",      1,  256,    128,                    []],
    [ "BO_",  "NmAsrMessage" ,              "Enumeration",  "", "",     "no",                   ["no","yes"]],
    [ "BU_",  "NmAsrCanMsgReducedTime" ,    "Integer",      1 , 65535,  320,                    []],
    [ "",     "NmAsrCanMsgCycleTime" ,      "Integer",      1,  65535,  640,                    []],
    [ "BU_",  "NmAsrCanMsgCycleOffset" ,    "Integer",      0,  65535,  0,                      []],
    [ "",     "NmAsrBaseAddress"            "Hex",          0x0, 0x7FF, 0x500,                  []],
    [ "BU_",  "ILUsed" ,                    "Enumeration",  "", "",     "no",                   ["no","yes"]],
    [ "",     "ILTxTimeout" ,               "Integer",      0,  65535,  0,                      []],
    [ "SG_",  "GenSigTimeoutValue" ,        "Integer",      0, 65535,   0,                      []],
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
    print(line_num)
    column_titles = sheet.row_values(0)
    index_list = get_column_index(DBC_FILE_COLUMN, column_titles)
    for i in range(1, line_num):
        row = sheet.row_values(i)
        dbc_data.append(row)
    print(dbc_data)
    return dbc_data

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
#     Method      : build_message_info
#
#     Description :
#           This method is build message information
#     Parameters  : data
#     Returns     : message line
# ===================================================================
def build_message_info(data):
    for info in data:
        if info[0] != "":
            print(info[0])
            print(data.index(info))

#====================function====================

#====================main========================
if __name__ == '__main__':
    print("main function")
    # data = read_matrix_file(r'C:\Users\Administrator\Desktop\data\dbc.xlsx')
    # build_message_info(data)
    create_dbc_file(r'C:\Users\kiwi\Desktop\data', 'T5')
#====================main========================