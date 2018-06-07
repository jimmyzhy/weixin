#!/usr/bin/env python
# -*- coding: utf-8 -*-

'''
from spec(.xls) to gererate testlink xml file
'''

import sys
import xlrd
reload(sys)
sys.setdefaultencoding("utf-8")
from xml.dom import minidom
import random


#### global value
EXTERNALID = 100
INTERNALID = 7654321
SUITE_NODE_ORDER = 0
CASE_NODE_OREDER = 0
DATA_MODEL = ''
ENTER = r'<br/>'
SPACE = r'&nbsp;'
P = r'<p>'
Pend = r'</p>'


steplist_r = [
    'Use GetParameterValues method get CPE default value',
    'Use SetParameterValues method set CPE value',
    'Use GetParameterValues method get CPE value should not change.',
    'Use SetParameterAttribute method change CPE node attibution.0:Notification-Non,1:Passive-Notification,2:Active-Notification',
    'Use GetParameterAttribute method get CPE node attibution.',
    ]

resultlist_r = [
    'CPE response default',
    'CPE response',
    'CPE response',
    'CPE response will not response any confirm information',
    'CPE response should same with attribute setting.'
    ]

steplist_w = [
    'Use GetParameterValues method get CPE default value',
    'Use SetParameterValues method set CPE value',
    'Use GetParameterValues method get CPE value should same with setting.',
    'Use SetParameterValues method set CPE invalid value',
    'Use GetParameterValues method get CPE value should not change.',
    'Use SetParameterAttribute method change CPE node attibution.0:Non-Notification,1:Passive-Notification,2:Active-Notification',
    'Use GetParameterAttribute method get CPE node attibution.'
    ]

resultlist_w = [
    'CPE response default',
    'CPE response',
    'CPE response',
    'CPE response',
    'CPE response',
    'CPE response will not response any confirm information',
    'CPE response should same with attribute setting.'
    ]

get_objective="Verify the data model node by TR069 RPC method ﻿SetParameterValues and GetParameterValues"
set_objective="Verify the data model node by TR069 RPC method  only support GetParameterValues"

## generate random 'string'
seed = '1234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ'
## generate random 'hexBinary'
seed_16 = '1234567890ABCDEFabcdef'

def rand(num,seed):
    list=[]
    for i in range(num):
        r = random.choice(seed)
        list.append(r)
    string=''.join(list)
    return string

## split the min, max number of int[], unsignedInt[]
def Int_Min_Max(datatype):
    # if datatype == 'int[-1:15][16:31]':
    #     valuelist = {'valid':[-1,0,15],'invalid':[max_+1,'abc']}
    # else:
    minnum = datatype.split('[')[1].split(']')[0].split(':')[0]
    maxnum = datatype.split('[')[1].split(']')[0].split(':')[1]
    if minnum != '':
        min_ = int(minnum)
    if maxnum != '':
        max_ = int(maxnum)
    if minnum == '':
        if datatype.startswith('int['):
            min_ = -5000
            if max_ > 0:
                valuelist = {'valid':[min_,0,max_],'invalid':[max_+1,'abc']}
            else:
                valuelist = {'valid':[min_,max_],'invalid':[max_+1,'abc']}
        elif datatype.startswith('unsignedInt['):
            min_ = 0
            valuelist = {'valid':[min_,max_],'invalid':[max_+1,'abc']}
    elif maxnum == '':
        max_ = 5000
        if min_ < 0:
            valuelist = {'valid':[min_,0,max_],'invalid':[min_-1,'abc']}
        else:
            valuelist = {'valid':[min_,max_],'invalid':[min_-1,'abc']}
    else:
        if min_ < 0 and max_ > 0:
            valuelist = {'valid':[min_,0,max_],'invalid':[min_-1,max_+1,'abc']}
        else:
            valuelist = {'valid':[min_,max_],'invalid':[min_-1,max_+1,'abc']}
    return valuelist

##
def String_Min_Max(datatype):
    if ':' in datatype:  ## string(a:b)
        minlen = datatype.split('(')[1].split(')')[0].split(':')[0]
        maxlen = datatype.split('(')[1].split(')')[0].split(':')[1]
        min_ = int(minlen)
        if min_ == 0:
            min_ = 1
        max_ = int(maxlen)
        valuelist = {'valid':[rand(min_,seed),rand(max_,seed)],'invalid':[rand(min_-1,seed),rand(max_+1,seed)]}
    else:    ## string(b)
        maxlen = datatype.split('(')[1].split(')')[0]
        max_ = int(maxlen)
        valuelist = {'valid':[rand(1,seed),rand(max_,seed)],'invalid':[rand(max_+1,seed)]}
    return valuelist


#################valuelist
def GetValueList(datatype):
    if datatype == 'boolean':
        valuelist = {'valid':['true','false'],'invalid':[2,'abc']}
    elif datatype == 'dataTime':
        valuelist = {'valid':['2011-08-22T00:55:54'],'invalid':['Thu Aug 22 19:24:28 2011']}
    elif datatype == 'int':
        valuelist = {'valid':[-32768,0,32767],'invalid':[123.123,'abc']}
    elif datatype.startswith('int['):
        valuelist = Int_Min_Max(datatype)
    elif datatype == 'unsignedInt':
        valuelist = {'valid':[0,65535],'invalid':[-1,'abc']}
    elif datatype.startswith('unsignedInt['):
        valuelist = Int_Min_Max(datatype)
    elif datatype == 'unsignedLong':
        valuelist = {'valid':[0,4294967295],'invalid':[-1,'abc']}
    elif datatype == 'long':
        valuelist = {'valid':[-2147483648,0,2147483647],'invalid':[-2147483649,2147483648,'abc']}
    elif datatype == 'string':
        valuelist = {'valid':['abcdefg','1234567890'],'invalid':[]}
    elif datatype.startswith('string('):
        valuelist = String_Min_Max(datatype)
    elif datatype.startswith('hexBinary('):
        min_ = 0
        max_ = 0
        if datatype == 'hexBinary(5:5)(13:13)':
            valuelist ={'valid':[rand(10, seed_16),rand(26, seed_16)],'invalid':[rand(8, seed_16),rand(12,seed_16),rand(24,seed_16),rand(28, seed_16)]}
        elif ':' in datatype:     ## hexBinary(a:b)
            minlen = datatype.split('(')[1].split(')')[0].split(':')[0]
            maxlen = datatype.split('(')[1].split(')')[0].split(':')[1]
            if minlen != '':
                min_ = int(minlen)
            if maxlen != '':
                max_ = int(maxlen)
            valuelist = {'valid':[rand(min_*2,seed_16),rand(max_*2,seed_16)],'invalid':[rand((min_-1)*2,seed_16),rand((max_+1)*2,seed_16),rand(max_*2-1,seed_16)]}
        else:    ## hexBinary(b)
            maxlen = datatype.split('(')[1].split(')')[0]
            max_ = int(maxlen)
            valuelist = {'valid':[rand(2,seed_16),rand(max_*2,seed_16)],'invalid':[rand(1,seed_16),rand((max_+1)*2,seed_16),rand(max_*2-1,seed_16)]}
    elif datatype.startswith('base64('):      ## =string
        valuelist = String_Min_Max(datatype)
    else:
        valuelist = {'valid':['abcdefg','1234567890'],'invalid':[]}  ## string
    return valuelist


#**************************************************
#####
def Summary(testcase,datatype,defaultvalue,validvalue,invalidvalue,access,samplevalue):

    case_summ_content = P + '<strong style="font-family: '+"'Trebuchet MS'"+', Verdana, Arial, sans-serif;">Objective:</strong>' + Pend
    case_summ_content = case_summ_content + P + 'Verify the data model node by TR069 RPC method'+SPACE+'﻿SetParameterValues and'+SPACE+'GetParameterValues'+SPACE + Pend
    case_summ_content = case_summ_content + '<p style="font-family: '+"'Trebuchet MS'"+', Verdana, Arial, sans-serif; background-color: rgb(238, 238, 238);">' + '<strong>Pre-condition:</strong>'+ Pend
    case_summ_content = case_summ_content + P +SPACE+ 'DUT in default mode' + Pend
    case_summ_content = case_summ_content + P + '<strong style="font-family: '+"'Trebuchet MS'"+', Verdana, Arial, sans-serif;">'+'Spec Parameters:'+'</strong>'+'<a id='+'"fck_paste_padding"'+'>﻿'+r'</a>' + Pend
    case_summ_content = case_summ_content + r'<table width="598" height="143" cellspacing="1" cellpadding="1" border="1">'
    if access == 'r':
        case_summ_content = case_summ_content + r'<tbody>'+r'<tr>'+r'<td>'+' Data Type'+r'</td>'+r'<td>'+SPACE+datatype+r'</td>'+r'</tr>'+ r'<tr>'+r'<td>'+' Default Value' + r'</td>' + r'<td>'+SPACE+defaultvalue+r'</td>' + r'</tr>'+r'<tr>'+r'<td>'+' Sample Value'+r'</td>'+r'<td>'+SPACE+samplevalue+r'</td>'+r'</tbody>'
    elif access == 'w':
        case_summ_content = case_summ_content + r'<tbody>'+r'<tr>'+r'<td>'+' Data Type'+r'</td>'+r'<td>'+SPACE+datatype+r'</td>'+r'</tr>'+ r'<tr>'+r'<td>'+' Default Value' + r'</td>' + r'<td>'+SPACE+defaultvalue+r'</td>' + r'</tr>'+r'<tr>'+r'<td>'+' Valid Value'+r'</td>'+r'<td>'+SPACE+validvalue+r'</td>'+r'<tr>'+r'<td>'+' Invalid Value'+r'</td>'+r'<td>'+SPACE+invalidvalue+r'</td>'+r'</tr>'+r'</tbody>'
    case_summ_content = case_summ_content + r'</table>'
    case_summ_content = case_summ_content + P + r'<a id="fck_paste_padding">'+SPACE+r'</a>'+ Pend
    return case_summ_content

#####
def Step(type,default,steplist,case_name,access):
    step_content=P
    if access == 'r':
        step1_content='1.'+SPACE*2+steplist[0]+ENTER
        step2_content='2.'+SPACE*2+steplist[1]+ENTER
        step2_content+=SPACE*3+r'&lt;'+'ParameterValueStruct'+r'&gt;'+ENTER
        step2_content+=SPACE*9+r'&lt;'+'Name'+r'&gt;'+case_name+r'&lt;'+r'/Name'+r'&gt;'+ENTER
        step2_content=step2_content+SPACE*9+r'&lt;'+'Value xsi:type='+r'&quot;'+'xsd:'+ type+r'&quot;'+r'&gt;'+'1234567890'+r'&lt;'+'/Value'+r'&gt;'+ENTER
        step2_content=step2_content+SPACE*7+r'&lt;'+'/ParameterValueStruct'+r'&gt;'+ENTER
        step3_content='3.'+SPACE*2+steplist[2]+ENTER
        step4_content='4.'+SPACE*2+steplist[3]+ENTER
        step4_content+=SPACE*3+r'&lt;'+'SetParameterAttributesStruct'+r'&gt;'+ENTER
        step4_content+=SPACE*9+r'&lt;'+'Name'+r'&gt;'+case_name+r'&lt;'+'/Name'+r'&gt;'+ENTER
        step4_content+=SPACE*9+r'&lt;'+'NotificationChange'+'&gt;'+'1'+r'&lt;'+'/NotificationChange'+r'&gt;'+ENTER
        step4_content+=SPACE*9+r'&lt;'+'Notification'+r'&gt;'+'0/1/2'+r'&lt;'+'/Notification'+r'&gt;'+ENTER
        step4_content+=SPACE*9+r'&lt;'+'AccessListChange'+r'&gt;'+'0'+r'&lt;'+'/AccessListChange'+r'&gt;'+ENTER
        step4_content+=SPACE*9+r'&lt;'+'AccessList SOAP-ENC:arrayType='+r'&quot;'+'xsd:string[0]'+r'&quot;'+'/'+r'&gt;'+ENTER
        step4_content+=SPACE*7+r'&lt;'+'/SetParameterAttributesStruct'+r'&gt;'+ENTER
        step5_content='5.'+SPACE*2+steplist[4]+ENTER
        step5_content+=SPACE*3+r'&lt;'+'cwmp:GetParameterAttributes xmlns:cwmp='+r'&quot;'+'urn:dslforum-org:cwmp-1-0'+r'&quot;'+r'&gt;'+ENTER
        step5_content+=SPACE*9+r'&lt;'+'ParameterNames SOAP-ENC:arrayType='+r'&quot;'+'xsd:string[1]'+r'&quot;'+r'&gt;'+ENTER
        step5_content+=SPACE*11+r'&lt;'+'string'+r'&gt;'+case_name+r'&lt;'+'/string'+r'&gt;'+ENTER
        step5_content+=SPACE*9+r'&lt;'+'/ParameterNames'+r'&gt;'+ENTER
        step5_content+=SPACE*7+r'&lt;'+'/cwmp:GetParameterAttributes'+r'&gt;'
        step_content+=step1_content+step2_content+step3_content+step4_content+step5_content+Pend
    elif access == 'w':
        step1_content = '1.'+SPACE*2+steplist[0]+ENTER
        step2_content = '2.'+SPACE*2+steplist[1]+ENTER+SPACE*3+r'&lt;'+'ParameterValueStruct'+r'&gt;'+ENTER+SPACE*9+r'&lt;'+'Name'+r'&gt;'+case_name+r'&lt;'+r'/Name'+r'&gt;'+ENTER
        step2_content = step2_content+SPACE*9+r'&lt;'+'Value xsi:type='+r'&quot;'+'xsd:'
        step2_content = step2_content+type+r'&quot;'+r'&gt;'+'Valid Value'+r'&lt;'+'/Value'+r'&gt;'+ENTER
        step2_content = step2_content+SPACE*7+r'&lt;'+'/ParameterValueStruct'+r'&gt;'+ENTER
        step3_content = '3.'+SPACE*2+steplist[2]+ENTER
        step4_content = '4.'+SPACE*2+steplist[3]+ENTER+SPACE*3+r'&lt;'+'ParameterValueStruct'+r'&gt;'+ENTER+SPACE*9+r'&lt;'+'Name'+r'&gt;'+case_name+r'&lt;'+r'/Name'+r'&gt;'+ENTER
        step4_content = step4_content+SPACE*9+r'&lt;'+'Value xsi:type='+r'&quot;'+'xsd:'
        step4_content = step4_content+type+r'&quot;'+r'&gt;'+'Invalid Value'+r'&lt;'+'/Value'+r'&gt;'+ENTER
        step4_content = step4_content+SPACE*7+r'&lt;'+'/ParameterValueStruct'+r'&gt;'+ENTER
        step5_content = '5.'+SPACE*2+steplist[4]+ENTER
        step6_content = '6.'+SPACE*2+steplist[5]+ENTER
        step6_content = step6_content+SPACE*9+r'&lt;'+'Name'+r'&gt;'+case_name+r'&lt;'+'/Name'+r'&gt;'+ENTER
        step6_content = step6_content+SPACE*9+r'&lt;'+'NotificationChange'+r'&gt;'+'1'+r'&lt;'+'/NotificationChange'+r'&gt;'+ENTER
        step6_content = step6_content+SPACE*9+r'&lt;'+'Notification'+r'&gt;'+'0/1/2'+r'&lt;'+'/Notification'+r'&gt;'+ENTER
        step6_content = step6_content+SPACE*9+r'&lt;'+'AccessListChange'+r'&gt;'+'0'+r'&lt;'+'/AccessListChange'+r'&gt;'+ENTER
        step6_content = step6_content+SPACE*9+r'&lt;'+'AccessList SOAP-ENC:arrayType='+r'&quot;'+'xsd:string[0]'+r'&quot;'+r'/&gt;'+ENTER
        step6_content = step6_content+SPACE*7+r'&lt;'+'/SetParameterAttributesStruct'+r'/&gt;'+ENTER
        step7_content = '7.'+SPACE*2+steplist[6]+ENTER
        step7_content = step7_content+SPACE*3+r'&lt;'+'cwmp:GetParameterAttributes xmlns:cwmp='+r'&quot;'+'urn:dslforum-org:cwmp-1-0'+r'&quot;'+r'&gt;'+ENTER
        step7_content = step7_content+SPACE*9+r'&lt;'+'ParameterNames SOAP-ENC:arrayType='+r'&quot;'+'xsd:string[1]'+r'&quot;'+r'&gt;'+ENTER
        step7_content = step7_content+SPACE*11+r'&lt;'+'string'+r'&gt;'+case_name+r'&lt;'+'/string'+r'&gt;'+ENTER
        step7_content = step7_content+SPACE*9+r'&lt;'+'/ParameterNames'+r'&gt;'+ENTER
        step7_content = step7_content+SPACE*7+r'&lt;'+'/cwmp:GetParameterAttributes'+r'&gt;'
        step_content += step1_content+step2_content+step3_content+step4_content+step5_content+step6_content+step7_content+Pend
    return step_content

######
def ExpectedResult(resultlist,case_name,default,type,access):
    result_content = P
    if access == 'r':
        result1_content='1.'+resultlist[0]+ENTER
        result1_content+=SPACE*3+r'&lt;'+'ParameterValueStruct'+r'&gt;'+ENTER
        result1_content+=SPACE*9+r'&lt;'+'Name'+r'&gt;'+case_name+r'&lt;'+'/Name'+r'&gt;'+ENTER
        result1_content+=SPACE*9+r'&lt;'+'Value xsi:type='+r'&quot;'+'xsd:'+type+r'&quot;'+r'&gt;'+default+r'&lt;'+'/Value'+r'&gt;'+ENTER
        result1_content+=SPACE*7+r'&lt;'+'/ParameterValueStruct'+r'&gt;'+ENTER
        result2_content='2.'+resultlist[1]+ENTER
        result2_content+=SPACE*3+r'&lt;'+'SOAP-ENV:Fault'+'&gt;'+ENTER
        result2_content+=SPACE*9+r'&lt;'+'faultcode'+r'&gt;'+'Client'+r'&lt;'+'/faultcode'+r'&gt;'+ENTER
        result2_content+=SPACE*9+r'&lt;'+'faultstring'+r'&gt;'+'CWMP fault'+r'&lt;'+'/faultstring'+r'&gt;'+ENTER
        result2_content+=SPACE*9+r'&lt;'+'detail'+r'&gt;'+ENTER
        result2_content+=SPACE*11+r'&lt;'+'cwmp:Fault'+r'&gt;'+ENTER
        result2_content+=SPACE*13+r'&lt;'+'FaultCode'+r'&gt;'+'9003'+r'&lt;'+'/FaultCode'+r'&gt;'+ENTER
        result2_content+=SPACE*13+r'&lt;'+'FaultString'+r'&gt;'+'Invalid arguments'+r'&lt;'+'/FaultString'+r'&gt;'+ENTER
        result2_content+=SPACE*13+r'&lt;'+'SetParameterValuesFault'+r'&gt;'+ENTER
        result2_content+=SPACE*15+r'&lt;'+'ParameterName'+r'&gt;'+case_name+r'&lt;'+'/ParameterName'+r'&gt;'+ENTER
        result2_content+=SPACE*15+r'&lt;'+'FaultCode'r'&gt;'+'9008'+r'&lt;'+'/FaultCode'+r'&gt;'+ENTER
        result2_content+=SPACE*15+r'&lt;'+'FaultString'+r'&gt;'+'Attempt to set a non-writable parameter'+r'&lt;'+'/FaultString'+r'&gt;'+ENTER
        result2_content+=SPACE*13+r'&lt;'+'/cwmp:Fault'+r'&gt;'+ENTER
        result2_content+=SPACE*9+r'&lt;'+'/detail'+r'&gt;'+ENTER
        result2_content+=SPACE*7+r'&lt;'+'/SOAP-ENV:Fault'+r'&gt;'+ENTER
        result3_content='3.'+resultlist[2]+ENTER
        result3_content+=SPACE*3+r'&lt;'+'ParameterValueStruct'+r'&gt;'+ENTER
        result3_content+=SPACE*9+r'&lt;'+'Name'+r'&gt;'+case_name+r'&lt;'+'/Name'+r'&gt;'+ENTER
        result3_content+=SPACE*9+r'&lt;'+'Value xsi:type='+r'&quot;'+r'xsd:'+type+r'&quot;'+r'&gt;'+default+r'&lt;'+'/Value'+r'&gt;'+ENTER
        result3_content+=SPACE*7+r'&lt;'+'/ParameterValueStruct'+r'&gt;'+ENTER
        result4_content='4.'+resultlist[3]+ENTER
        result4_content+=SPACE*3+r'&lt;'+'cwmp:SetParameterAttributesResponse/'+r'&gt;'+ENTER
        result5_content='5.'+resultlist[4]+ENTER
        result5_content+=SPACE*3+r'&lt;'+'cwmp:GetParameterAttributesResponse'+r'&gt;'+ENTER
        result5_content+=SPACE*5+r'&lt;'+'ParameterList SOAP-ENC:arrayType='+r'&quot;'+'ParameterAttributeStruct[1]'+r'&quot;'+r'&gt;'+ENTER
        result5_content+=SPACE*7+r'&lt;'+'ParameterAttributeStruct'+r'&gt;'+ENTER
        result5_content+=SPACE*9+r'&lt;'+'Name'+r'&gt;'+case_name+r'&lt;'+'/Name'+r'&gt;'+ENTER
        result5_content+=SPACE*9+r'&lt;'+'Notification'+r'&gt;'+'0/1/2'+r'&lt;'+'/Notification'+r'&gt;'+ENTER
        result5_content+=SPACE*9+r'&lt;'+'AccessList xsi:type='+r'&quot;'+'tns:AccessList'+r'&quot;'+r'/&gt;'+ENTER
        result5_content+=SPACE*7+r'&lt;'+'/ParameterAttributeStruct'+r'&gt;'+ENTER
        result5_content+=SPACE*5+r'&lt;'+'/ParameterList'+r'&gt;'+ENTER
        result5_content+=SPACE*3+r'&lt;'+'/cwmp:GetParameterAttributesResponse'+r'&gt;'
        result_content+=result1_content+result2_content+result3_content+result4_content+result5_content+Pend
    elif access == 'w':
        result1_content = '1.'+resultlist[0]+ENTER
        result1_content += SPACE*3+r'&lt;'+'ParameterValueStruct'+r'&gt;'+ENTER
        result1_content += SPACE*9+r'&lt;'+'Name'+r'&gt;'+case_name+r'&lt;'+'/Name'+r'&gt;'+ENTER
        result1_content += SPACE*9+r'&lt;'+'Value xsi:type='+r'&quot;'+'xsd:'
        result1_content = result1_content+type+r'&quot;'+r'&gt;'+default+r'&lt;'+'/Value'+r'&gt;'+ENTER
        result1_content += SPACE*7+r'&lt;'+'/ParameterValueStruct'+r'&gt;'+ENTER
        result2_content = '2.'+resultlist[1]+ENTER
        result2_content += SPACE*3+r'&lt;'+'cwmp:SetParameterValuesResponse'+r'&gt;'+ENTER
        result2_content += SPACE*9+r'&lt;'+'Status xsi:type='+r'&quot;'+'xsd:int'+r'&quot;'+r'&gt;'+'0'+r'&lt;'+'/Status'+r'&gt;'+ENTER
        result2_content += SPACE*6+r'&lt;'+'/cwmp:SetParameterValuesResponse'+r'&gt;'+ENTER
        result3_content = '3.'+resultlist[2]+ENTER
        result3_content += SPACE*3+r'&lt;'+'ParameterValueStruct'+r'&gt;'+ENTER
        result3_content += SPACE*9+r'&lt;'+'Name'+r'&gt;'+case_name+r'&lt;'+'/Name'+r'&gt;'+ENTER
        result3_content += SPACE*9+r'&lt;'+'Value xsi:type='+r'&quot;'+r'xsd:'+type+r'&quot;'+r'&gt;'+'Valid Value'+r'&lt;'+'/Value'+r'&gt;'+ENTER
        result3_content += SPACE*7+r'&lt;'+'/ParameterValueStruct'+r'&gt;'+ENTER
        result4_content = '4.'+resultlist[3]+ENTER
        result4_content += SPACE*3+r'&lt;'+'SOAP-ENV:Fault'+r'&gt;'+ENTER
        result4_content += SPACE*9+r'&lt;'+'faultcode'+r'&gt;'+'Client'+r'&lt;'+'/faultcode'+r'&gt;'+ENTER
        result4_content += SPACE*9+r'&lt;'+'faultstring'+r'&gt;'+'CWMP fault'+r'&lt;'+'/faultstring'+r'&gt;'+ENTER
        result4_content += SPACE*9+r'&lt;'+'detail'r'&gt;'+ENTER
        result4_content += SPACE*11+r'&lt;'+'cwmp:Fault'+r'&gt;'+ENTER
        result4_content += SPACE*13+r'&lt;'+'FaultCode'+r'&gt;'+'9003'+r'&lt;'+'/FaultCode'+r'&gt;'+ENTER
        result4_content += SPACE*13+r'&lt;'+'FaultString'+r'&gt;'+'Invalid arguments'+r'&lt;'+'/FaultString'+r'&gt;'+ENTER
        result4_content += SPACE*13+r'&lt;'+'SetParameterValuesFault'+r'&gt;'+ENTER
        result4_content += SPACE*15+r'&lt;'+'ParameterName'+r'&gt;'+case_name+r'&lt;'+'/ParameterName'+r'&gt;'+ENTER
        result4_content += SPACE*15+r'&lt;'+'FaultCode'+r'&gt;'+'9007'+r'&lt;'+'/FaultCode'+r'&gt;'+ENTER
        result4_content += SPACE*15+r'&lt;'+'FaultString'+r'&gt;'+'Invalid Parameter value'+r'&lt;'+'/FaultString'+r'&gt;'+ENTER
        result4_content += SPACE*13+r'&lt;'+'/SetParameterValuesFault'+r'&gt;'+ENTER
        result4_content += SPACE*11+r'&lt;'+'/cwmp:Fault'+r'&gt;'+ENTER
        result4_content += SPACE*9+r'&lt;'+'/detail'+r'&gt;'+ENTER
        result4_content += SPACE*7+r'&lt;'+'/SOAP-ENV:Fault'+r'&gt;'+ENTER
        result5_content = '5.'+resultlist[4]+ENTER
        result5_content += SPACE*3+r'&lt;'+'ParameterValueStruct'+r'&gt;'+ENTER
        result5_content += SPACE*9+r'&lt;'+'Name'+r'&gt;'+case_name+r'&lt;'+'/Name'+r'&gt;'+ENTER
        result5_content += SPACE*9+r'&lt;'+'Value xsi:type='+r'&quot;'+'xsd:'+type+r'&quot;'+r'&gt;'+'Vaild Value [-1]'+r'&lt;'+'/Value'+r'&gt;'+ENTER
        result5_content += SPACE*7+r'&lt;'+'/ParameterValueStruct'+r'&gt;'+ENTER
        result6_content = '6.'+resultlist[5]+ENTER
        result6_content += SPACE*3+r'&lt;'+'cwmp:SetParameterAttributesResponse/'+r'&gt;'+ENTER
        result7_content = '7.'+resultlist[6]+ENTER
        result7_content += SPACE*3+r'&lt;'+'cwmp:GetParameterAttributesResponse'+r'&gt;'+ENTER
        result7_content += SPACE*5+r'&lt;'+'ParameterList SOAP-ENC:arrayType='+r'&quot;'+'ParameterAttributeStruct[1]'+r'&quot;'+r'&gt;'+ENTER
        result7_content += SPACE*7+r'&lt;'+'ParameterAttributeStruct'+r'&gt;'+ENTER
        result7_content += SPACE*9+r'&lt;'+'Name'+r'&gt;'+case_name+r'&lt;'+'/Name'+r'&gt;'+ENTER
        result7_content += SPACE*9+r'&lt;'+'Notification'+r'&gt;'+'0/1/2'+r'&lt;'+'/Notification'+r'&gt;'+ENTER
        result7_content += SPACE*9+r'&lt;'+'AccessList xsi:type='+r'&quot;'+'tns:AccessList'+r'&quot;'+r'/&gt;'+ENTER
        result7_content += SPACE*7+r'&lt;'+'/ParameterAttributeStruct'+'&gt;'+ENTER
        result7_content += SPACE*5+r'&lt;'+'/ParameterList'+r'&gt;'+ENTER
        result7_content += SPACE*3+r'&lt;'+'/cwmp:GetParameterAttributesResponse'+r'&gt;'
        result_content += result1_content+result2_content+result3_content+result4_content+result5_content+result6_content+result7_content+Pend
    return result_content

#####



###############################################
def CreateOneCase(doc,list,Attribute):

    global EXTERNALID
    global INTERNALID
    global SUITE_NODE_ORDER
    global CASE_NODE_OREDER
    global DATA_MODEL

    support = list[0]    # Support Y/N
    nodename = list[1]

    datatype = list[2]
    valuelist = GetValueList(datatype)
    if 'string' in datatype:
        type='string'
    elif 'dateTime' in datatype:
        type='datetime'
    elif 'boolean' in datatype:
        type='boolean'
    elif 'unsigned' in datatype:
        type='unsignedInt'
    elif 'int' in datatype or 'long' in datatype:
        type='int'
    elif 'signed' in datatype:
        type='signedint'
    else:
        type='string'   # other datatype

    access = list[3]
    # description = list[4]

    defaultvalue = list[5]
    if defaultvalue == '':
        defaultvalue = '-'

    default = defaultvalue
    if default == '-':
        default = ''

    samplevalue = list[6]    ## Read-Only parameter
    if samplevalue == '':
        samplevalue = '-'    ## Read-Write parameter
    validvalue = list[6]
    if validvalue == '':
        validvalue = valuelist['valid']
        for i in range(0,len(validvalue)):
            validvalue[i] = str(validvalue[i])
        validvalue = "; ".join(validvalue)

    invalidvalue = valuelist['invalid']
    for i in range(0,len(invalidvalue)):
        invalidvalue[i] = str(invalidvalue[i])
    invalidvalue = "; ".join(invalidvalue)
    if invalidvalue == '':
        invalidvalue = '-'


    #***********************************************#
    if support == 'N':
        return 0
    elif support == 'Y':
        if datatype == 'object':
            DATA_MODEL = nodename.strip()
            return 0
        else:
            if access == 'r':
                if Attribute == 'w':
                    return 0
                elif Attribute == 'r':
                    steplist = steplist_r
                    resultlist = resultlist_r
            elif access == 'w':
                if Attribute == 'r':
                    return 0
                elif Attribute == 'w':
                    steplist = steplist_w
                    resultlist = resultlist_w

            case_name = DATA_MODEL + nodename.strip()
            testcase = doc.createElement('testcase')
            testcase.setAttribute('name',case_name)
            INTERNALID +=1
            testcase.setAttribute('internalid',str(INTERNALID))
            CASE_NODE_OREDER +=1
            test_node_order = doc.createElement('node_order')
            test_node_order.appendChild(doc.createCDATASection(str(CASE_NODE_OREDER)))
            testcase.appendChild(test_node_order)
            EXTERNALID +=1
            test_extid = doc.createElement('externalid')
            test_extid.appendChild( doc.createCDATASection(str(EXTERNALID)))
            testcase.appendChild(test_extid)

            test_summ = doc.createElement('summary')
            case_summ_content = Summary(testcase,datatype,defaultvalue,validvalue,invalidvalue,access,samplevalue)
            test_summ.appendChild(doc.createCDATASection(case_summ_content))
            testcase.appendChild(test_summ)

            test_step = doc.createElement('steps')
            step_content = Step(type,defaultvalue,steplist,case_name,access)
            test_step.appendChild(doc.createCDATASection(step_content))
            testcase.appendChild(test_step)

            test_expect = doc.createElement('expectedresults')
            result_content = ExpectedResult(resultlist,case_name,default,type,access)
            test_expect.appendChild(doc.createCDATASection(result_content))
            testcase.appendChild(test_expect)

            return testcase


#############################################################
def main():

    #*********running parameter********#
    helpnote = 'The script requires 6 parameters:\n'
    helpnote += '   para1: ExcelFile  e.g. TR069_DataModles_Definition.xls\n'
    helpnote += '   para2: SheetFile  e.g. tr181\n'
    helpnote += '   para3: Attribute  e.g. r or w\n'
    helpnote += '          -- node is writable or non-writable\n'
    helpnote += '   para4: XmlFile    e.g. tr181-r-1.xml\n'
    helpnote += '          -- Attention: the size of XmlFile must less than 8MB\n'
    helpnote += '          -- When generating big file, we should set "StartRow","EndRow" to different values to divide into several parts\n'
    helpnote += '   para5: StartRow   e.g. 1\n'
    helpnote += '          -- start row of excel for current generation\n'
    helpnote += '   para6: EndRow     e.g. 1000\n'
    helpnote += '          -- end row of excel for current generation\n'
    helpnote += '   -- e.g. part1: 1 1000, part2: 1000 2000, part3: 2000 3314\n'
    helpnote += '   para7: SuitName   e.g. TR181-R-Part1\n'

    if sys.argv[1] == '-h':
        print helpnote
    elif len(sys.argv) != 8:    ## + sys.argv[0]
        print '<ERROR> Wrong Parametr Number!\r\n' + helpnote
    else:
        ExcelFile = sys.argv[1]
        SheetFile = sys.argv[2]
        Attribute = sys.argv[3]
        XmlFile = sys.argv[4]
        StartRow = int(sys.argv[5])
        EndRow = int(sys.argv[6])
        SuitName = sys.argv[7]


        doc = minidom.Document()
        testsuite = doc.createElement('testsuite')
        testsuite.setAttribute('name',SuitName)
        doc.appendChild(testsuite)
        node_order = doc.createElement('node_order')
        node_order.appendChild(doc.createCDATASection('1'))
        details = doc.createElement('details')
        details.appendChild(doc.createCDATASection(''))
        testsuite.appendChild(node_order)
        testsuite.appendChild(details)


        ## read excel
        book = xlrd.open_workbook(ExcelFile)
        sheet = book.sheet_by_name(SheetFile)

        nrows = sheet.nrows   ## 获得行数
        ncols = sheet.ncols   ## 获得列数

        # List_table = []
        for i in range(StartRow,EndRow):
            List = []    ## list combined with a line data
            for j in range(0,ncols):
                cellvalue = sheet.cell(i,j).value
                List.append(cellvalue)
            # print List[0]
            # List_table.append(List)
            testcase = CreateOneCase(doc,List,Attribute)
            if testcase != 0:
                testsuite.appendChild(testcase)

        fileID = open(XmlFile,'w')
        fileID.write(doc.toprettyxml())
        fileID.flush()
        fileID.close()


###############################################
if __name__ == "__main__":
    main()
