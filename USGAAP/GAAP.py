# coding = utf-8
# python 3.7.0

'''
符合美国现行会计准则的收入计算模型的一些配置
包括环境配置和业务配置
'''

import os
import csv
import xlrd
import time
import datetime
import multiprocessing
from math import floor
from platform import system
from numpy import pmt, irr
from codecs import open as op
from calendar import monthrange


###############################################系统环境参数##############################################################
"""程序名称"""
PROGRAM = 'US GAAP REVENUE PROVISION MODEL'
#当前系统环境
SYSTEM = system()
#当前系统时间
TIME = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
#获取今天的日期
TODAY = datetime.date.today().strftime("%Y/%m/%d")
#获取昨天的日期
YESTERDAY = (datetime.date.today() - datetime.timedelta(days=1)).strftime("%Y/%m/%d")
#获取上个月时间
LAST_MONTH = (datetime.date(datetime.date.today().year, datetime.date.today().month, 1)
              - datetime.timedelta(days = 1)).strftime("%Y-%m")
#统计时间
STATIS = (datetime.date.today() - datetime.timedelta(days=1)).strftime("%Y%m%d")
#设置时间范围
BEGIN = '2014/12/01'
"""
备注：全部日期统一的格式为'YYYY/MM/DD'
"""

#################################################全局配置################################################################

'''
定义全局参数配置(可通过hive进行配置，配置表可见 dashu_dspmopr.fin_usgaap_config， 通过set_config函数进行配置)
'''

#程序参数配置
config = {
    'PROCESS': 4,                            #启动进程数量
    'HEADER': 1,                             #输出数据包含表头：1 不包含：0
    'MONTH': 0,                              #输出还款计划表 按月：1 按日：0
    'OUTPUT_SCHEDULE': 0,                    #输出还款计划表：1 不输出：0
    'OUTPUT_REPORT': 1,                      #输出收入报表：1 不输出：0
    'COMBINE': 1,                            #是否合并输出结果：1 不输出：0
    'UPLOAD': 0,                             #是否上传数据：1 不上传：0
    'OUTPUT_SCHEDULE_FILE': 'schedule',      #还款计划表文件命名
    'OUTPUT_REPORT_FILE': 'report',          #收入报表文件命名
}

#################################################定义字段################################################################

"""定义字段（可修改调整）"""
last_end = 'last_end'       #上一期间末
start = 'start'             #开始时间
end = 'end'                 #结束时间
putout_no = 'putout_no'     #出账编号
sortno = 'sortno'           #排序编号
total = 'total'             #总和
avg = 'avg'                 #平均
putout_date = 'putout_date' #放款日期
putout_end = 'putout_end'   #还款计划最后一天
actual_day_num = 'actual_day_num' #当前期间实际天数
accumulate_day_num = 'accumulate_day_num' #放款日到当前期间的实际累加天数
date = 'date'   #日期
day_num = 'day_num'  #天数
ctype = 'ctype' #合作机构
lineid = 'lineid' #合作机构代码
term = 'term'   #贷款期限(年)
loan_term = 'loan_term'  #贷款期限(月)
business_sum = 'business_sum'  #放款金额
principal = 'principal'   #剩余本金
discount_rate = 'discount_rate'   #折现率
interest_rate = 'interest_rate'  #利息率
present_value_rate = 'present_value_rate' #现值比率
interest_to_bank = 'interest_to_bank'   #偿还银行利息
service_rate = 'service_rate' #服务费率
service_fee = 'service_fee'   #服务费
service_fee_discount = 'service_fee_discount'   #服务费折现
service_fee_accumulate = 'service_fee_accumulate'   #服务费累加
management_rate = 'management_rate'   #月管理费率
management_rate1 = 'management_rate1' #月管理费率——子账户1
management_rate2 = 'management_rate2' #月管理费率——子账户2
management_fee = 'management_fee'       #月管理费
management_fee1 = 'management_fee1'     #月管理费——子账户1
management_fee2 = 'management_fee2'     #月管理费——子账户2
cash_match_rate = 'cash_match_rate'   #现金撮合收入系数
guarantee_ratio = 'guarantee_ratio'   #担保比例
margin_ratio = 'margin_ratio' #全周期利润率
loss_ratio = 'loss_ratio'  #全周期损失率

future_loss_ratio = 'future_loss_ratio' #未来损失率
future_loss_month = 'future_loss_month' #未来损失月份
pre_future_loss_ratio = 'pre_future_loss_ratio' #以前未来损失率
pre_future_loss_month = 'pre_future_loss_month' #以前未来损失月份
premium_receivable_day1 = 'premium_receivable_day1' #取自还款计划表第0天的premium receivable
contract_assets_day1 = 'contract_assets_day1'       #取自还款计划表第0天的contract assets
premium_receivable_loss = 'premium_receivable_loss' #premium receivable-day1 * 未来损失率
contract_assets_loss = 'contract_assets_loss'       #contract asset-day1 * 未来损失率
pre_premium_receivable_loss = 'pre_premium_receivable_loss' #premium receivable-day1 * 未来损失率
pre_contract_assets_loss = 'pre_contract_assets_loss'       #contract asset-day1 * 未来损失率


guarantee_loss = 'guarantee_loss' #担保收入损失
pre_guarantee_loss = 'pre_guarantee_loss' #以前期间担保收入损失
active = 'active'
non_active = 'non_active'
finish_date = 'finish_date' #结清日
overdue_date = 'overdue_date' #逾期日
fix_overdue_date = 'fix_overdue_date' #逾期日计息截止日 = 逾期日 + 30天
overdue_days = 'overdue_days' #逾期天数
mob = "mob" #month on book
buy_date = 'buy_date' #代偿日
buy_sum = 'buy_sum' #代偿金额
ASC450 = 'ASC450'  #预计赔付金额
loan_status = 'loan_status'  #贷款状态
actual_end_date = 'actual_end_date' #计息截止日
pre_actual_end_date = 'pre_actual_end_date' #以前计息截止日
active_status = 'active_status' #存量状态
unearned_premium_reserve = 'unearned_premium_reserve'    #未到期责任准备金
unearned_premium_reserve_amorize = 'unearned_premium_reserve_amorize' #未到期责任准备金-摊销数
guarantee_liability_460 = 'guarantee_liability_460' #未到期责任准备金
pre_cash_for_guarantee = 'pre_cash_for_guarantee'                       #以前期间现金-担保
pre_guarantee = 'pre_guarantee'                                         #以前期间担保
pre_interest_adjust_for_guarantee = 'pre_interest_adjust_for_guarantee' #以前期间利息调整-担保
pre_post_origination = 'pre_post_origination'                           #以前期间贷后服务
pre_cash_for_matching = 'pre_cash_for_matching'                         #以前期间现金-撮合服务
pre_matching = 'pre_matching'                                           #以前期间撮合服务
pre_match_early_repay = 'pre_match_early_repay'                         #以前期间撮合服务（提前还款）
pre_interest_adjust_for_match_early_repay = 'pre_interest_adjust_for_match_early_repay' #以前期间利息调整-撮合服务（提前还款）
pre_cash_for_post_origination = 'pre_cash_for_post_origination'         #以前期间现金-贷后服务
impairment_loss1 = 'impairment_loss1'                                   #资产减值损失-应收保费-代偿
pre_impairment_loss1 = 'pre_impairment_loss1'                           #以前期间资产减值损失-应收保费-代偿
impairment_loss2 = 'impairment_loss2'                                   #资产减值损失-合同资产-代偿
pre_impairment_loss2 = 'pre_impairment_loss2'                           #以前期间资产减值损失-合同资产-代偿'
fair_value_guarantees = 'fair_value_guarantees'  #公允价值-担保
pre_fair_value_guarantees = 'pre_fair_value_guarantees'  #以前期间公允价值-担保
charge_income = 'charge_income'  #计入利润表拨备
pre_charge_income = 'pre_charge_income'  #以前期间计入利润表拨备
carry = 'carry'  #结转
adjust_status = 'adjust_status'   #调整状态
itemname = 'itemname'  #一级产品
total_guarantee = 'total_guarantee' #总担保收入
upfront_cost = 'upfront_cost'   #贷前成本
on_going_cost = 'on_going_cost'   #贷后成本
total_cost = 'total_cost'   #总成本
cash_flow_to_bank = 'cash_flow_to_bank'   #现金流(银行)
management_fee_discount = 'management_fee_discount'   #月管理费(折现)
cash_flow_of_client = 'cash_flow_of_client'   #借款人现金流出
contract_value = 'contract_value'   #合同价值
contract_value_discount = 'contract_value_discount'   #合同价值(折现)
matching = 'matching'   #撮合服务
match_early_repay = 'match_early_repay'   #撮合服务（提前还款）
match_early_rate = 'match_early_rate'             #撮合服务收入比例（repayment时费用减少，一般只收取 撮合收入 * 0.94）
post_origination = 'post_origination'   #贷后服务
guarantee = 'guarantee'   #担保
revenue_margin = 'revenue_margin'    #收入-利润
cash_for_matching = 'cash_for_matching'   #现金-撮合服务
cash_for_matching1 = 'cash_for_matching1' #现金-撮合服务 - 子账户
cash_for_post_origination = 'cash_for_post_origination'   #现金-贷后服务
cash_for_guarantee = 'cash_for_guarantee'   #现金-担保
premium_receivable_principal = 'premium_receivable_principal'   #应交保费本金
premium_receivable_principal_amorize = 'premium_receivable_principal_amorize'   #应交保费本金-摊销数
contract_assets = 'contract_assets'   #合同资产-撮合服务
contract_assets1 = 'contract_assets1' #合同资产-撮合服务 - 子账户
contract_assets_early_repay = 'contract_assets_early_repay'   #合同资产-撮合服务(提前还款)
contract_assets_early_repay_amorize ='contract_assets_early_repay_amorize'   #合同资产-撮合服务(提前还款)-摊销数
interest_adjust_for_match = 'interest_adjust_for_match'   #利息调整-撮合服务
interest_adjust_for_guarantee = 'interest_adjust_for_guarantee'   #利息调整-担保
interest_adjust_for_match_early_repay = 'interest_adjust_for_match_early_repay'   #利息调整-撮合服务（提前还款）
created_by = 'created_by'   #创建人
created_date = 'created_date'   #创建时间
updated_by = 'updated_by'   #更新人
update_date = 'update_date' #更新时间

"""定义中文字段"""
unadjusted = 'unadjusted'  #调整前
adjust_num = 'adjust_num'  #调整
adjusted = 'adjusted'      #调整后
compensatory = 'compensatory'   #代偿结清
repayment =   'repayment'                 #提前还款
accrued_loan =  'accrued_loan'                #应计贷款
non_accrued_loan = 'non_accrued_loan'   #非应计贷款
dashu = 'dashu'    #大数时贷

'''参数字段'''
UPFRONT_COST = 'UPFRONT_COST'
ON_GOING_COST = 'ON_GOING_COST'
MATCH_EARLY_RATE = 'MATCH_EARLY_RATE'
TPY_Model_ratio1 = 'TPY_Model_ratio1'
TPY_Model_ratio2 = 'TPY_Model_ratio2'
TPY_Model_ratio3 = 'TPY_Model_ratio3'
TPY_Model_ratio4 = 'TPY_Model_ratio4'

'''新增字段'''
month_of_loan = 'month_of_loan'     #截止日月数
implied_price_concession_ratio = 'implied_price_concession_ratio'   #隐含价格特许权比率
reduction_of_interest_guarantee = 'reduction_of_interest_guarantee' #减息担保
reduction_of_interest_matching = 'reduction_of_interest_matching'   #减息撮合

month_of_loan_end = 'month_of_loan_end'
implied_price_concession_ratio_end = 'implied_price_concession_ratio_end'
reduction_of_interest_guarantee_accumulate = 'reduction_of_interest_guarantee_accumulate'
reduction_of_interest_matching_accumulate = 'reduction_of_interest_matching_accumulate'

month_of_loan_start = 'month_of_loan_start'
implied_price_concession_ratio_start = 'implied_price_concession_ratio_start'
reduction_of_interest_guarantee_current = 'reduction_of_interest_guarantee_current'
reduction_of_interest_matching_current = 'reduction_of_interest_matching_current'

#################################################财务函数类##############################################################

"""财务函数类，可单独调用函数进行使用"""
class Math(object):
    def __init__(self, precision=8):
        """初始化计算结果保留位数，默认保留6位数（四舍五入）"""
        self.precision = precision

    @staticmethod
    def get_value(value):
        """获取数值"""
        if value == '' or value is None:
            return 0
        else:
            return float(value.replace(',', ''))            #实际赔付金额

    @staticmethod
    def date_sub(date_1, date_2):
        """计算日期之间的天数
           比较日期大小 YYYY/MM/DD格式
        """
        date_1 = datetime.datetime.strptime(date_1, "%Y/%m/%d")
        date_2 = datetime.datetime.strptime(date_2, "%Y/%m/%d")
        return (date_1-date_2).days

    @staticmethod
    def get_day_num(term):
        """放款天数 = 放款期限 X 365 + 1"""
        return int(term * 365 + 1)

    @staticmethod
    def get_day_num_monthly(current_date):
        """当月天数 = 日期（"%Y-%m-%d"/"%Y/%m/%d"字符串格式） 的 当月天数"""
        return monthrange(int(current_date[0:4]), int(current_date[5:7]))[1]

    @staticmethod
    def get_is_include(value, values):
        """判断是否在字符串内"""
        for v in values:
            if value[0:len(v)] == v:
                return True
        return False

    @staticmethod
    def get_date(current_date, n):
        current_date = datetime.datetime.strptime(current_date, '%Y/%m/%d')
        """日期 = 放款日 + 第n天"""
        return (current_date + datetime.timedelta(days=n)).strftime('%Y/%m/%d')

    def set_precision(self, precision):
        """设置精度（四舍五入）"""
        self.precision = precision

    def __to_precision(self, value):
        """保留精度（四舍五入）"""
        return round(value, self.precision)

    def accumulate_value(self, last_value, value):
        """数值累加 = 上一个数值 + 当前数值"""
        return self.__to_precision(last_value + value)


    def subtract_value(self, last_value, value):
        """数值累加 = 上一个数值 - 当前数值"""
        return self.__to_precision(last_value - value)

    def multiply_value(self,last_value,value):
        """数值相乘 = 上一个数值 X 当前数值"""
        return self.__to_precision(last_value * value)

    def divide_value(self, value1, value2):
        """平均值 = 数值 / 数量n"""
        return self.__to_precision(value1 / value2)

    def get_irr_value(self, principal, service_rate, management_rate, interest_rate, term):
        """内部收益率 = IRR(借款人未来每月现金流出)"""
        cash_flow = []
        period = 12
        month_day = 30
        month_num = 0  #默认第几个月收取服务费
        cash_flow_month = principal * management_rate + pmt(interest_rate / period, term * period, -principal, 0, 0)

        if self.get_is_include(self.p[lineid], config['SERVER'].keys()):
            cash_flow.append(0 - principal)
            """计算是第几个月收取服务费"""
            month_num = int(round(float(config['SERVER'][self.p[lineid]])/30, 0))
        else:
            cash_flow.append(principal * service_rate - principal)

        term = int(term)
        for i in range(term * period):
            if i == month_num-1:
                cash_flow.append(cash_flow_month + principal * service_rate)
            else:
                cash_flow.append(cash_flow_month)
        return self.__to_precision(irr(cash_flow) / month_day)

    def get_service_fee(self, principal, service_rate):
        """服务费 = 本金 X 服务费率"""
        return self.__to_precision(principal*service_rate)

    def get_guarantee(self, principal, guarantee_ratio):
        """担保费 = 本金 X 担保比例"""
        return self.__to_precision(principal*guarantee_ratio)

    def get_revenue_margin(self, principal, margin_ratio, day_num):
        """每日收入-利润 = 本金 X 全周期利润率 / 总天数"""
        return self.__to_precision(principal*margin_ratio/day_num)

    def get_present_value_rate(self, discount_rate, n_day):
        """现值比率 = 1/（1+折现率）的n次方"""
        return self.__to_precision(pow(1 / (1 + discount_rate), n_day))

    def get_cash_flow_to_bank(self,interest_rate,term,daynum,principal):
        """每日现金流(银行）= PMT( 利率 X 期限 / 总天数,总天数,-本金,0,0)"""
        return self.__to_precision(pmt(interest_rate * term / daynum, daynum, -principal, 0, 0))

    def get_cash_flow_of_client(self,cash_flow_bank_day, management_fee, service_fee=0, principal=0):
        """借款人现金流出 = 每日现金流(银行）+ 每日管理费用 + 服务费（放款当日计算） - 本金（放款当日计算）"""
        return self.__to_precision(cash_flow_bank_day + management_fee + service_fee - principal)

    def get_interest_to_bank(self,principal, interest_rate):                     #计算银行利息
        """每日银行利息 = 本金 X 利率 / 365天"""
        return self.__to_precision(principal * interest_rate / 365)

    def get_principal_rest(self,last_principal, cash_flow_bank_day, bank_interest):    #计算剩余本金
        """剩余本金 = 前一天剩余本金 - 当日现金流(银行）+ 当日银行利息"""
        return self.__to_precision(last_principal - cash_flow_bank_day + bank_interest)

    def get_present_value(self, value, present_value_rate):
        """现值 = 数值 X 现值比例
        """
        return self.__to_precision(value * present_value_rate)

    def get_management_fee(self, principal, management_rate, date):  #计算每天管理费
        """当日管理费 = 本金 X 月管理费率 / 当月天数
        """
        management_fee = principal * management_rate / self.get_day_num_monthly(date)
        return self.__to_precision(management_fee)

    def get_match_service(self, contract_value_discount, guarantee, upfront_cost, match_rate=1):
        """撮合收入 = （合同折现价值 - 担保收入）X 贷前成本比例 X 比例（若repayment）"""
        return self.__to_precision((contract_value_discount - guarantee) * upfront_cost * match_rate)

    def get_post_origination(self, contract_value_discount, guarantee, on_going_cost):
        """贷后服务收入 = （合同折现价值 - 担保收入）X 贷后成本比例"""
        return self.__to_precision((contract_value_discount - guarantee) * on_going_cost)

    def get_cash_for_match_service(self, matching, service_discount, total_contract_value_discount, management_fee):
        """现金-撮合服务收入 =  每天管理费用 X（撮合收入 - 服务费折现）/ （合同折现价值 - 服务费折现）"""
        return self.__to_precision(management_fee *
                                   (matching - service_discount) /
                                   (total_contract_value_discount - service_discount))

    def get_cash_for_post_origination(self, contract_value_discount, service_fee, management_fee, post_origination):
        """现金-贷后服务收入 =  每天管理费用 X 贷后服务收入 / （合同折现价值 - 服务费折现）"""
        return self.__to_precision(management_fee * post_origination/(contract_value_discount-service_fee))

    def get_cash_for_guarantee(self, contract_value_discount, service_fee, management_fee, guarantee):
        """现金-担保收入 =  每天合同价值 X 担保收入 / （合同折现价值 - 服务费折现）"""
        return self.__to_precision(management_fee * guarantee / (contract_value_discount - service_fee))

    def get_premium_receivable_principal(self, last_premium_receivable_principal, cash_for_guarantee, discount_rate):
        """应收保费本金 =  前一天应收保费本金 X （1+折现率）- 现金-担保收入"""
        return self.__to_precision(last_premium_receivable_principal * (1 + discount_rate) - cash_for_guarantee)

    def get_contract_assets(self, last_contact_assets, discount_rate, cash_for_match_service):
        """合同资产-撮合服务 = 前一天合同资产-撮合服务 X （1+折现率）- 现金-撮合服务收入"""
        return self.__to_precision(last_contact_assets * (1 + discount_rate) - cash_for_match_service)

    def get_contract_assets_early_repayment(self, start_contact_assets, last_contact_assets, discount_rate,
                                            cash_for_match_service):
        """合同资产-撮合服务-repayment的计算逻辑如下：

            如果前一天合同资产-撮合服务-repayment > 0
            则合同资产-撮合服务-repayment = MAX （前一天合同资产-撮合服务-repayment X （1+折现率）- 现金-撮合服务收入, 0）
            如果前一天合同资产-撮合服务-repayment <= 0
            则合同资产-撮合服务-repayment = MIN （前一天合同资产-撮合服务-repayment X （1+折现率）- 现金-撮合服务收入, 0）
        """
        if start_contact_assets > 0:
            return max(self.get_contract_assets(last_contact_assets, discount_rate, cash_for_match_service), 0)
        else:
            return min(self.get_contract_assets(last_contact_assets, discount_rate, cash_for_match_service), 0)

    def get_interest_adjustment_for_matching_early_repayment(self, last_contract_assets_early_repayment,
                                                                    contract_assets_early_repayment,
                                                                    discount_rate,
                                                                    cash_for_match_service):
        """利息调整 - 撮合服务 - repayment的计算逻辑如下：

            如果合同资产-撮合服务-repayment = 0
                如果 前一天合同资产-撮合服务-repayment >0
                    则利息调整-撮合服务-repayment = MIN （撮合服务收入 - 前一天合同资产-撮合服务-repayment, 撮合服务收入）
                否则  利息调整-撮合服务-repayment = MAX （撮合服务收入 - 前一天合同资产-撮合服务-repayment, 撮合服务收入）
            如果合同资产-撮合服务-repayment != 0
            则利息调整-撮合服务-repayment = 前一天合同资产-撮合服务-repayment X 折现率
        """
        if contract_assets_early_repayment == 0:
            if last_contract_assets_early_repayment>0:
                return min(cash_for_match_service-last_contract_assets_early_repayment, cash_for_match_service)
            else:
                return max(cash_for_match_service-last_contract_assets_early_repayment, cash_for_match_service)
        else:
            return self.multiply_value(last_contract_assets_early_repayment,discount_rate)

################################################收入模型计算#############################################################

class Revenue(Math):
    """贷款收入模型类,继承数学模型类中的运算函数"""
    def __init__(self, config, param, period=None, future_loss=None, changeList=None, ipcr=None):
        '''
        传入参数
            config: 全局变量
            param: 某一笔贷款数据，例如: ['RL20151210000075', '廊坊银行', 'RL20150105000002', 36.0, 100000.0, 2015/12/11, 9.0,
                                        3.0, 0.0, 0.0, 0.59, '2018/12/11', '', '', 0.09210366435822859, 0.06973581093255772, 0.02505]
            period: 计算区间, 例如：{'2014-12-2018-12': {'start': '2014/12/01', 'end': '2018/12/31'}}
            future_loss: DataEngine类中读取config配置表第二页得到的参数
            changeList: DataEngine类中读取config配置表第一页和第五页组合得到的参数, 业务参数及其修改历史
            ipcr: DataEngine类中读取config配置表第三页得到的参数
        '''
        '''继承父类属性'''
        super(Revenue, self).__init__()                                         #继承父类的构造方法，如数值精度

        self.future_loss = future_loss
        self.ipcr = ipcr

        '''初始化参数，创建参数词典'''
        self.p = {}
        self.p[itemname] = dashu
        self.p[putout_no] = param[0]                                                        #出账编号
        self.p[ctype] = param[1]                                                            #合作机构
        self.p[lineid] = param[2]
        self.p[loan_term] = int(param[3])                                                   #贷款期限（月)
        self.p[term] = round(self.p[loan_term]/12, 0)                                       #贷款期限（年)
        self.p[business_sum] = self.p[principal] = self.get_value(str(param[4]))            #发放金额
        self.p[putout_date] = param[5]                                                      #还款计划表日期
        self.p[interest_rate] = float(param[6]) / 100                                       #折现率
        self.p[service_rate] = float(param[7]) / 100                                        #服务费率

        #判断浮动管理费分母是否为0
        if str(param[9]) not in ['0.0', '0', '']:
            #如果不为0，则管理费率1 = 浮动管理费分子 / 分母
            self.p[management_rate1] = float(param[8]) / float(param[9])                    #管理费1
        else:
            #如果为0，则管理费率1为0
            self.p[management_rate1] = 0
        self.p[management_rate] = self.p[management_rate2] = float(param[10]) / 100         #管理费率2，显示出来

        #判断机构id是否是太保模式，若机构是太保模式，担保比例使用guarantee_ratio_taibao
        if self.get_is_include(self.p[lineid], config['FILTER1']):
            self.p[guarantee_ratio] = round(float(param[-2]), 6)                                #担保比例
        else:
            self.p[guarantee_ratio] = round(float(param[-3]), 6)

        self.p[margin_ratio] = float(param[-1])                                             #全周期利润率

        #loss_ratio = guarantee_ratio - margin
        self.p[loss_ratio] = self.subtract_value(self.p[guarantee_ratio], self.p[margin_ratio])

        self.p[discount_rate] = self.get_irr_value(self.p[principal], self.p[service_rate], self.p[management_rate2],
                                                   self.p[interest_rate], self.p[term])     #折现率
        self.p[finish_date] = param[11]                                                     #结清日
        self.p[overdue_date] = param[12]                                                    #逾期日
        self.p[buy_date] = param[13]                                                        #代偿日

        #业务参数列表
        bus_param = [UPFRONT_COST, ON_GOING_COST, MATCH_EARLY_RATE,
                     TPY_Model_ratio1, TPY_Model_ratio2,
                     TPY_Model_ratio3, TPY_Model_ratio4]
        #字符串转化为日期格式
        strp_date = datetime.datetime.strptime(param[5], '%Y/%m/%d')
        #配置与出账日期相对的业务参数
        for i in range(0, len(changeList)):
            for key in sorted(changeList[i].keys(), reverse=True):
                if strp_date >= key:
                    config[bus_param[i]] = changeList[i][key]
                    break

        '''保存全局变量，保持一致'''
        self.p[upfront_cost] = config['UPFRONT_COST']
        self.p[on_going_cost] = config['ON_GOING_COST']
        self.p[match_early_rate] = config['MATCH_EARLY_RATE']
        self.p[TPY_Model_ratio1] = config['TPY_Model_ratio1']
        self.p[TPY_Model_ratio2] = config['TPY_Model_ratio2']
        self.p[TPY_Model_ratio3] = config['TPY_Model_ratio3']
        self.p[TPY_Model_ratio4] = config['TPY_Model_ratio4']

        '''基础数据初始化计算'''
        #计算该笔贷款服务费
        self.p[service_fee] = self.get_service_fee(self.p[principal], self.p[service_rate])
        #计算该笔贷款担保费
        self.p[total_guarantee] = self.get_guarantee(self.p[principal], self.p[guarantee_ratio])
        #计算该笔贷款总天数
        self.p[day_num] = self.get_day_num(self.p[term])
        #计算该笔贷款收入-利润
        self.p[revenue_margin] = self.get_revenue_margin(self.p[principal], self.p[margin_ratio], self.p[day_num])

        #还款计划表初始化
        self.date_list = []
        self.schedule = {}
        self.schedule_init()
        #报表收入计算时间段范围
        self.period_list = period

    '''还款计划表初始化'''
    def schedule_init(self):
            self.schedule[avg] = {}
            self.schedule[total] = {}

            #从放款日到最后一天进行初始化
            for n_day in range(0, self.p[day_num] + 1):
                now = self.get_date(self.p[putout_date], n_day)
                self.date_list.append(now)
                self.schedule[now] = {}
            #放款日最后一天
            self.p[putout_end] = self.date_list[-1]

    '''第一层：现金流计算程序'''
    def cash_flow_calculation(self):
        #放款日当天数据初始化
        start = self.date_list[0]
        #每日现金流(银行）
        self.schedule[avg][cash_flow_to_bank] = self.get_cash_flow_to_bank(self.p[interest_rate], self.p[term],
                                                                           self.p[day_num], self.p[principal])

        #第一天现值比例为1
        self.schedule[start][present_value_rate] = 1
        if self.get_is_include(self.p[lineid], config['SERVER'].keys()):
            #判定是第几天收取服务费
            self.schedule[self.get_date(start, config['SERVER'][self.p[lineid]])][service_fee] = self.p[service_fee]
            self.schedule[start][service_fee] = 0
        else:
            #放款日当天收取服务费
            self.p[service_fee_discount] = self.p[service_fee]
            self.schedule[start][service_fee] = self.p[service_fee]

        self.schedule[start][service_fee_accumulate] = self.schedule[start][service_fee]

        self.schedule[start][principal] = self.p[principal]

        #借款人现金流出初始化
        self.schedule[start][cash_flow_of_client] = self.get_cash_flow_of_client(0, 0, self.schedule[start][service_fee],
                                                                                 self.p[principal])
        #合同价值（折现）初始化
        self.schedule[start][contract_value] = self.schedule[start][service_fee]
        self.schedule[start][contract_value_discount] = self.schedule[start][service_fee]
        self.schedule[total][contract_value] = self.schedule[start][service_fee]
        self.schedule[total][contract_value_discount] = self.schedule[start][service_fee]

        for n in range(1, self.p[day_num] + 1):
            #从放款日第二天开始循环计算
            last = self.date_list[n - 1]    #前一天日期
            now = self.date_list[n]         #当天日期

            #现值比率
            self.schedule[now][present_value_rate] = self.get_present_value_rate(self.p[discount_rate], n)
            if service_fee not in self.schedule[now].keys():
                self.schedule[now][service_fee] = 0
            else:
                self.p[service_fee_discount] = self.schedule[now][service_fee] * self.schedule[now][present_value_rate]

            self.schedule[now][service_fee_accumulate] = self.accumulate_value(self.schedule[last][service_fee_accumulate],
                                                                               self.schedule[now][service_fee])
            #现金流（银行）
            self.schedule[now][cash_flow_to_bank] = self.schedule[avg][cash_flow_to_bank]
            #利息（银行）
            self.schedule[now][interest_to_bank] = self.get_interest_to_bank(self.schedule[last][principal],
                                                                             self.p[interest_rate])
            #剩余本金
            self.schedule[now][principal] = self.get_principal_rest(self.schedule[last][principal],
                                                                    self.schedule[avg][cash_flow_to_bank],
                                                                    self.schedule[now][interest_to_bank])
            #管理费用1 根据银行利息浮动分润
            self.schedule[now][management_fee1] = self.multiply_value(self.schedule[now][interest_to_bank],
                                                                      self.p[management_rate1])
            #管理费用2 固定的收取方式
            if self.p[lineid] == 'RL20180103360499':
                if self.date_sub(now, self.p[putout_date]) <= 365:
                    self.schedule[now][management_fee2] = self.get_management_fee(self.p[principal], self.p[management_rate], now)
                else:
                    self.schedule[now][management_fee2] = 0
            else:
                self.schedule[now][management_fee2] = self.get_management_fee(self.p[principal], self.p[management_rate], now)

            #管理费用 合并上面的费用
            self.schedule[now][management_fee] = self.schedule[now][management_fee1] + self.schedule[now][management_fee2]
            self.schedule[now][management_fee_discount] = self.get_present_value(self.schedule[now][management_fee],
                                                                                 self.schedule[now][present_value_rate])
            #借款人现金流出
            self.schedule[now][cash_flow_of_client] = self.get_cash_flow_of_client(self.schedule[avg][cash_flow_to_bank],
                                                                                   self.schedule[now][management_fee2],
                                                                                   self.schedule[now][service_fee])
            #合同价值
            self.schedule[now][contract_value] = self.schedule[now][service_fee] + self.schedule[now][management_fee]
            #合同价值-折现
            self.schedule[now][contract_value_discount] = self.get_present_value(self.schedule[now][contract_value],
                                                                                 self.schedule[now][present_value_rate])
            #合同价值总和
            self.schedule[total][contract_value] = self.accumulate_value(self.schedule[total][contract_value],
                                                                         self.schedule[now][contract_value])
            #合同价值折现总和
            self.schedule[total][contract_value_discount] = self.accumulate_value(self.schedule[total][contract_value_discount],
                                                                                  self.schedule[now][contract_value_discount])

    '''第二层：收入计算程序'''
    def revenue_calculation(self):
        #放款日当天数据初始化
        start = self.date_list[0]

        #撮合收入
        self.schedule[start][matching] = self.get_match_service(self.schedule[total][contract_value_discount],
                                                                self.p[total_guarantee], config['UPFRONT_COST'])
        #撮合收入（提前还款）
        self.schedule[start][match_early_repay] = self.subtract_value(self.schedule[start][matching],
                                                                      self.schedule[total][contract_value_discount] *
                                                                      (1 - config['MATCH_EARLY_RATE']))
        #贷后服务收入
        self.schedule[total][post_origination] = self.get_post_origination(self.schedule[total][contract_value_discount],
                                                                           self.p[total_guarantee], config['ON_GOING_COST'])
        #担保费用
        self.schedule[total][guarantee] = self.p[total_guarantee]
        #平均每天担保费用
        self.schedule[avg][guarantee] = self.divide_value(self.schedule[total][guarantee], self.p[day_num])
        #现金-撮合收入 放款日 等于合同价值
        self.schedule[start][cash_for_matching] = self.schedule[start][contract_value_discount]
        #应收保费本金 放款日 等于担保收入
        self.schedule[start][premium_receivable_principal] = self.p[total_guarantee]
        #合同资产-撮合服务  初始化
        self.schedule[start][contract_assets1] = 0

        for n in range(1, self.p[day_num]+1):
            #从放款日第二天开始循环计算
            last = self.date_list[n-1]   #前一天日期
            now = self.date_list[n]      #当天日期

            #平均每天担保费用
            self.schedule[now][guarantee] = self.schedule[avg][guarantee]
            #平均每天利润-收入
            self.schedule[now][revenue_margin] = self.p[revenue_margin]
            #现金-撮合服务
            self.schedule[now][cash_for_matching1] = self.get_cash_for_match_service(self.schedule[start][matching],
                                                                                     self.p[service_fee_discount],
                                                                                     self.schedule[total][contract_value_discount],
                                                                                     self.schedule[now][management_fee])

            self.schedule[now][cash_for_matching] = self.schedule[now][cash_for_matching1] + self.schedule[now][service_fee]

            #合同资产-撮合服务  累加 每天 现金-撮合服务 X 现值比例
            self.schedule[start][contract_assets1] = self.accumulate_value(self.schedule[start][contract_assets1],
                                                                           self.schedule[now][cash_for_matching1] *
                                                                           self.schedule[now][present_value_rate])
            #现金-贷后服务
            self.schedule[now][cash_for_post_origination] = self.get_cash_for_post_origination(self.schedule[total][contract_value_discount],
                                                                                               self.p[service_fee_discount],
                                                                                               self.schedule[now][management_fee],
                                                                                               self.schedule[total][post_origination])
            #贷后服务每日收入
            self.schedule[now][post_origination] = self.schedule[now][cash_for_post_origination]
            #现金-担保
            self.schedule[now][cash_for_guarantee] = self.get_cash_for_guarantee(self.schedule[total][contract_value_discount],
                                                                                 self.p[service_fee_discount],
                                                                                 self.schedule[now][management_fee],
                                                                                 self.schedule[total][guarantee])
            #应收保费本金
            self.schedule[now][premium_receivable_principal] = self.get_premium_receivable_principal(self.schedule[last][premium_receivable_principal],
                                                                                                     self.schedule[now][cash_for_guarantee],
                                                                                                     self.p[discount_rate])
            #应收保费本金-摊销数
            self.schedule[now][premium_receivable_principal_amorize] = self.subtract_value(self.schedule[last][premium_receivable_principal],
                                                                                           self.schedule[now][premium_receivable_principal])

    '''第三层：资产负债表计算程序'''
    def balance_sheet_calculation(self):
        #放款日当天数据初始化
        start = self.date_list[0]

        self.schedule[start][contract_assets] = self.schedule[start][contract_assets1] + (self.p[service_fee]-self.schedule[start][service_fee])
        #合同资产-撮合服务(repayment) 初始化
        self.schedule[start][contract_assets_early_repay] = self.subtract_value(self.schedule[start][match_early_repay],
                                                                                self.schedule[start][service_fee])
        for n in range(1, self.p[day_num] + 1):
            #从放款日第二天开始循环计算
            last = self.date_list[n-1]   #前一天日期
            now = self.date_list[n]      #当天日期

            #合同资产-撮合服务
            self.schedule[now][contract_assets1] = self.get_contract_assets(self.schedule[last][contract_assets1],
                                                                            self.p[discount_rate],
                                                                            self.schedule[now][cash_for_matching1])
            self.schedule[now][contract_assets] = self.schedule[now][contract_assets1] + self.p[service_fee] - \
                                                  self.schedule[now][service_fee_accumulate]
            #合同资产-撮合服务(repayment)
            self.schedule[now][contract_assets_early_repay] = self.get_contract_assets_early_repayment(
                                                                   self.schedule[start][contract_assets_early_repay],
                                                                   self.schedule[last][contract_assets_early_repay],
                                                                   self.p[discount_rate],
                                                                   self.schedule[now][cash_for_matching])
            #合同资产-撮合服务(repayment)-摊销数
            self.schedule[now][contract_assets_early_repay_amorize] = self.subtract_value(self.schedule[last][contract_assets_early_repay],
                                                                                          self.schedule[now][contract_assets_early_repay])
            #利息调整-撮合服务
            self.schedule[now][interest_adjust_for_match] = self.multiply_value(self.schedule[last][contract_assets],
                                                                                self.p[discount_rate])
            #利息调整-撮合服务（提前还款）
            self.schedule[now][interest_adjust_for_match_early_repay] = self.multiply_value(self.schedule[last][contract_assets],
                                                                                            self.p[discount_rate])
            #利息调整-担保服务
            self.schedule[now][interest_adjust_for_guarantee] = self.multiply_value(self.schedule[last][premium_receivable_principal],
                                                                                    self.p[discount_rate])

    '''
        定义输出字段，可以对需要输出的字段进行调整修改
        h1 为参数字段
        h2 为还款计划表计算输出字段
        h3 记录创建人和时间
    '''
    def output_schedule_list(self, writer):
        #h1、h2、h3为还款计划表的表头
        h1 = [putout_no, putout_date, sortno, loan_term, date, start, end,                  # 必要字段
              interest_rate, guarantee_ratio, margin_ratio, service_rate, management_rate,  # 导入参数
              discount_rate, upfront_cost, on_going_cost, total_guarantee]                  # 中间计算生成参数

        h2 = [principal, present_value_rate, cash_flow_to_bank, interest_to_bank,
              service_fee, management_fee, management_fee_discount, cash_flow_of_client, contract_value,
              contract_value_discount, matching,
              premium_receivable_principal,
              premium_receivable_principal_amorize, contract_assets, contract_assets_early_repay,
              contract_assets_early_repay_amorize, interest_adjust_for_match,
              match_early_repay, post_origination, guarantee,
              interest_adjust_for_guarantee, interest_adjust_for_match_early_repay,
              cash_for_matching, cash_for_post_origination, cash_for_guarantee]

        h3 = [created_by, created_date]

        #合并表头字段
        head = []
        head.extend(h1)
        head.extend(h2)
        head.extend(h3)

        if config['HEADER'] == 1:
            #是否输出表头
            writer.writerow(head)
        table = []
        sortno_num = 0

        #输出按月为单位的还款计划表
        if config['MONTH'] == 1:
            #取时间点值的字段
            point = [principal, present_value_rate, premium_receivable_principal, contract_assets, contract_assets_early_repay]
            #取时间段累加值的字段
            sum = [cash_flow_to_bank, interest_to_bank,
                   service_fee, management_fee, management_fee_discount, cash_flow_of_client, contract_value,
                   contract_value_discount, matching, premium_receivable_principal_amorize,
                   contract_assets_early_repay_amorize, interest_adjust_for_match,
                   match_early_repay, post_origination, guarantee, revenue_margin,
                   interest_adjust_for_guarantee, interest_adjust_for_match_early_repay,
                   cash_for_matching, cash_for_post_origination, cash_for_guarantee]
            n = len(self.date_list)
            self.row = {}

            #循环每一天进行压缩计算
            for i in range(0, n):
                dt = self.date_list[i]
                month = dt[0:4] + '-' + dt[5:7]
                month_start = dt[0:4] + '/' + dt[5:7] + '/01'
                month_end = dt[0:4] + '/' + dt[5:7] + '/' + str(self.get_day_num_monthly(dt))
                if month not in self.row.keys():
                    self.row[month] = {}
                if i + 1 < n:
                    next_month = self.date_list[i + 1][0:4] + '-' + self.date_list[i + 1][5:7]
                else:
                    next_month = None

                #循环取累加值的字段
                for key2 in sum:
                    if key2 in self.schedule[dt].keys():
                        if key2 in self.row[month].keys():
                            self.row[month][key2] = self.accumulate_value(self.row[month][key2], self.schedule[dt][key2])
                        else:
                            self.row[month][key2] = self.schedule[dt][key2]

                if month != next_month:
                    #循环h1字段
                    for key1 in h1:
                        if key1 == sortno:
                            self.row[month][key1] = sortno_num
                            sortno_num = sortno_num + 1
                        elif key1 == date:
                            self.row[month][key1] = month
                        elif key1 == start:
                            self.row[month][key1] = month_start
                        elif key1 == end:
                            self.row[month][key1] = month_end
                        else:
                            self.row[month][key1] = self.p[key1]
                    #循环取时间点数据的字段
                    for key2 in point:
                        self.row[month][key2] = self.schedule[dt][key2]
                    #循环h3字段
                    for key3 in h3:
                        if key3 == created_by:
                            self.row[month][key3] = PROGRAM
                        elif key3 == created_date:
                            self.row[month][key3] = TIME
                    line = []
                    for key in head:
                        if key not in self.row[month].keys():
                            line.append(0)
                        else:
                            line.append(self.row[month][key])
                    table.append(line)
        else:
            #输出按天的还款计划表
            for dt in self.date_list:
                line = []
                #循环h1字段
                for key1 in h1:
                    if key1 == sortno:
                        line.append(sortno_num)
                        sortno_num = sortno_num + 1
                    elif key1 == date:
                        line.append(dt)
                    elif key1 == start:
                        line.append(dt)
                    elif key1 == end:
                        line.append(dt)
                    else:
                        line.append(self.p[key1])
                #循环h2字段
                for key2 in h2:
                    if key2 not in self.schedule[dt].keys():
                        line.append(0)
                    else:
                        line.append(self.schedule[dt][key2])
                #循环h3字段
                for key3 in h3:
                    if key3 == created_by:
                        line.append(PROGRAM)
                    elif key3 == created_date:
                        line.append(TIME)

                table.append(line)
        #表格数据写入文本
        writer.writerows(table)

    '''收入报表相关函数'''
    def __to_precision(self, value):
        #保留精度（四舍五入）
        return round(value, self.precision)

    '''获取贷款状态'''
    def get_loan_status(self):
        if self.p[buy_date] != '':
            if self.date_sub(self.p[buy_date], self.p[end]) <= 0:
                return compensatory, self.p[buy_date]
        if self.p[finish_date] != '':
            if self.date_sub(self.p[finish_date], self.p[end]) <= 0:
                return repayment, self.p[finish_date]
        if self.p[overdue_date] != '':
            if self.date_sub(self.p[overdue_date], self.p[end]) <= 0:
                return non_accrued_loan,self.p[overdue_date]
            else:
                return non_accrued_loan, self.p[end]
        return accrued_loan, self.p[end]

    '''获取代偿日,结清日，逾期日天，期间结束日期 最早的一个作为计息截止日'''
    def get_real_actual_end_date(self):
        for dt in [self.p[buy_date], self.p[finish_date], self.p[overdue_date], self.p[end]]:
            if dt != '':
                if self.date_sub(self.p[actual_end_date], dt) >= 0:
                    self.p[actual_end_date] = dt

    '''获取上一个计息截止日'''
    def get_pre_actual_date(self):
        if self.p[actual_end_date] != '':
            if self.date_sub(self.p[actual_end_date], self.p[last_end]) <= 0:
                return self.p[actual_end_date]
            else:
                return self.p[last_end]

    '''判断该笔贷款是否是存量贷款'''
    def get_active_status(self):
        if self.date_sub(self.p[putout_date], self.p[end]) <= 0:
            self.p[active_status] = active
        else:
            self.p[active_status] = non_active

    '''计算累加值'''
    def get_accumulate_value(self, start_date, end_date, account):
        n = self.date_sub(end_date, start_date)+1
        value = 0
        for i in range(n):
            current_date = self.get_date(start_date, i)
            if current_date in self.schedule.keys():
                if account in self.schedule[current_date].keys():
                    value = value + self.schedule[current_date][account]
        return self.__to_precision(value)

    '''计算多个累加值'''
    def get_accumulate_value_list(self, start_date, end_date, account_list):
        n = self.date_sub(end_date, start_date)+1
        value = {}
        for account in account_list:
            value[account] = 0
        for i in range(n):
            current_date = self.get_date(start_date, i)
            if current_date in self.schedule.keys():
                for account in account_list:
                    if account in self.schedule[current_date].keys():
                        if account in value.keys():
                            value[account] = value[account] + self.schedule[current_date][account]
        for account in account_list:
            value[account] = self.__to_precision(value[account])
        return value

    '''获取某个时间点的值'''
    def get_time_point_value(self, current_date, account):

        if current_date not in self.schedule.keys():
            return 0
        return self.schedule[current_date][account]

    '''判断是否需要进行结转  1表示需要结转，0表示不需要结转'''
    def get_carry_value(self):
        if self.p[loan_status] == repayment:
            #如果是提前还款，根据结清日期判断是否结转
            if self.date_sub(self.p[finish_date],self.p[start]) < 0:
                return 1

        if self.p[loan_status] == compensatory:
            #如果是代偿结清，根据代偿日期判断是否结转
            if self.date_sub(self.p[buy_date],self.p[start]) < 0:
                return 1
        return 0

    '''计算MOB'''
    def get_mob(self):
        if self.p[overdue_date] == '':
            self.p[overdue_days] = 0
            self.p[mob] = 0
        else:
            self.p[overdue_days] = self.date_sub(TODAY, self.p[overdue_date])
            self.p[mob] = int(self.p[overdue_days] / 31)+1
            if self.p[mob] > 13:
                self.p[mob] = 13

    '''判断是否在时间段date1和date2之间'''
    def get_include_month(self, date0, date1, date2):
        if self.date_sub(date0, date1) < 0:
            return False
        if self.date_sub(date0, date2) > 0:
            return False
        return True

    '''收入报表初始化'''
    def report_init(self):
        self.report = {}
        for period in self.period_list.keys():
            #期间数据初始化
            self.report[period] = {}
            self.report[period][start] = self.period_list[period][start]
            self.report[period][end] = self.period_list[period][end]

            #期间以前数据初始化
            self.report[period + '_before'] = {}
            self.report[period + '_before'][start] = BEGIN
            self.report[period + '_before'][end] = self.get_date(self.period_list[period][start], -1)

    def get_future_loss_month(self, date):
        import math
        #期间结束日 - 放款日>=o，取未来损失率的值
        num = self.date_sub(date, self.p[putout_date])
        if num >= 0:
            #【(期间结束日 - 放款日) / 30 向上取整数 再加 12】 跟 【贷款期限】 取 小
            return min(math.ceil(float(num)/30) + 12, int(self.p[loan_term]))
        else:
            return 0

    def get_future_loss_ratio(self, month):
        if month == 0:
            return 0
        if month > 36:
            month = 36

        for quater in self.future_loss.keys():

            check = self.get_include_month(self.p[putout_date],
                                           self.future_loss[quater][start],
                                           self.future_loss[quater][end])
            if check is True:
                return self.future_loss[quater][month]

    '''
        调整数
        提前还款和代偿结清的贷款进行以下的会计处理
    '''
    def carry_value(self, period=None):

        if self.p[loan_status] in [repayment, compensatory]:
            for account in [contract_assets_early_repay, premium_receivable_principal, guarantee_liability_460]:
                self.report[period][adjust_num][account] = -self.report[period][unadjusted][account]

                self.report[period][adjusted][account] = \
                    self.report[period][unadjusted][account] + self.report[period][adjust_num][account]
            #是否进行结转操作
            if self.p[carry] == 0:
                #不结转
                self.report[period][adjust_num][impairment_loss1] = \
                    self.report[period][unadjusted][premium_receivable_principal]
                self.report[period][adjust_num][impairment_loss2] = \
                    self.report[period][unadjusted][contract_assets_early_repay]

                for account in [impairment_loss1, impairment_loss2]:
                    self.report[period][adjusted][account] = self.report[period][adjust_num][account]

                #guarantee 调整数 等同于 guarantee_liability_460
                self.report[period][adjust_num][guarantee] = \
                    self.report[period][unadjusted][guarantee_liability_460]

                #guarantee 进行调整
                self.report[period][adjusted][guarantee] = \
                    self.report[period][unadjusted][guarantee] + \
                    self.report[period][adjust_num][guarantee]
            else:
                #结转
                self.report[period][adjust_num][pre_impairment_loss1] = \
                    self.report[period][unadjusted][premium_receivable_principal]
                self.report[period][adjust_num][pre_impairment_loss2] = \
                    self.report[period][unadjusted][contract_assets_early_repay]

                for account in [pre_impairment_loss1, pre_impairment_loss2]:
                    self.report[period][adjusted][account] = self.report[period][adjust_num][account]

                #pre_guarantee 调整数 等同于 guarantee_liability_460
                self.report[period][adjust_num][pre_guarantee] = \
                    self.report[period][unadjusted][guarantee_liability_460]

                #pre_guarantee 进行调整
                self.report[period][adjusted][pre_guarantee] = \
                    self.report[period][unadjusted][pre_guarantee] + \
                    self.report[period][adjust_num][pre_guarantee]

    '''到截止日的贷款月数'''
    @staticmethod
    def get_month_of_loan(date, putout_date):
        date = datetime.datetime.strptime(date, '%Y/%m/%d')
        putout_date = datetime.datetime.strptime(putout_date, '%Y/%m/%d')
        days = (date - putout_date).days
        if days > 0:
            month = int(days) / 30
            month = floor(month)
            if month < 36:
                return month
            else:
                return 36
        else:
            return 0

    '''计算减息担保'''
    @staticmethod
    def get_reduction_of_interest_guarantee(loss_ratio, premium_receivable_day1, ipcr):
        return round(float(loss_ratio) * float(premium_receivable_day1) * ipcr, 6)

    '''计算减息撮合'''
    def get_reduction_of_interest_matching(self, loss_ratio, contract_assets_day1, ipcr):
        return round(loss_ratio * contract_assets_day1 * ipcr, 6)


    '''收入报表计算程序'''
    def report_calculation(self):
        self.report_init()
        #计算mob和逾期天数
        self.get_mob()

        #循环每个期间并开始进行报表计算
        for period in self.period_list.keys():
            self.p[last_end] = self.report[period + '_before'][end]     #上个期间末
            self.p[start] = self.report[period][start]                  #期初
            self.p[end] = self.report[period][end]                      #期末

            self.get_active_status()
            self.report[period][active_status] = self.p[active_status]

            #未来损失率新的计算逻辑
            self.report[period][future_loss_month] = self.get_future_loss_month(self.p[end])
            self.report[period][pre_future_loss_month] = self.get_future_loss_month(self.p[last_end])
            self.report[period][future_loss_ratio] = self.get_future_loss_ratio(self.report[period][future_loss_month])
            self.report[period][pre_future_loss_ratio] = self.get_future_loss_ratio(self.report[period][pre_future_loss_month])
            self.report[period][premium_receivable_day1] = self.get_time_point_value(self.p[putout_date], premium_receivable_principal)
            self.report[period][contract_assets_day1] = self.get_time_point_value(self.p[putout_date], contract_assets_early_repay)

            self.report[period][month_of_loan_end] = self.get_month_of_loan(self.report[period][end], self.p[putout_date])
            # 通过month_of_loan匹配ipcr值
            if self.report[period][month_of_loan_end] in self.ipcr.keys():
                self.report[period][implied_price_concession_ratio_end] = self.ipcr[self.report[period][month_of_loan_end]]
            else:
                self.report[period][implied_price_concession_ratio_end] = 0
            # reduction_of_interest_guarantee_accumulate = loss_ratio * premium_receivable_day1 * implied_price_concession_ratio_end
            self.report[period][reduction_of_interest_guarantee_accumulate] = self.get_reduction_of_interest_guarantee(
                                                                                  self.p[loss_ratio],
                                                                                  self.report[period][premium_receivable_day1],
                                                                                  self.report[period][implied_price_concession_ratio_end])
            # reduction_of_interest_matching_accumulate = loss_ratio * contract_assets_day1 * implied_price_concession_ratio_end
            self.report[period][reduction_of_interest_matching_accumulate] = self.get_reduction_of_interest_matching(
                                                                                  self.p[loss_ratio],
                                                                                  self.report[period][contract_assets_day1],
                                                                                  self.report[period][implied_price_concession_ratio_end])
            # start - putout_date
            self.report[period][month_of_loan_start] = self.get_month_of_loan(self.report[period][start],
                                                                              self.p[putout_date])
            #取与month_of_loan_start对应的ipcr值
            if self.report[period][month_of_loan_start] in self.ipcr.keys():
                self.report[period][implied_price_concession_ratio_start] = self.ipcr[self.report[period][month_of_loan_start]]
            else:
                self.report[period][implied_price_concession_ratio_start] = 0

            #reduction_of_interest_guarantee_current = loss_ratio * premium_receivable_day1 * （implied_price_concession_ratio_end - implied_price_concession_ratio_start）
            self.report[period][reduction_of_interest_guarantee_current] = self.get_reduction_of_interest_guarantee(
                        self.p[loss_ratio], self.report[period][premium_receivable_day1],
                        self.report[period][implied_price_concession_ratio_end]-self.report[period][implied_price_concession_ratio_start])

            #reduction_of_interest_matching_current = loss_ratio * contract_assets_day1 * （implied_price_concession_ratio_end - implied_price_concession_ratio_start）
            self.report[period][reduction_of_interest_matching_current] = self.get_reduction_of_interest_matching(
                        self.p[loss_ratio], self.report[period][contract_assets_day1],
                        self.report[period][implied_price_concession_ratio_end]-self.report[period][implied_price_concession_ratio_start])

            #获取贷款状态，计算时间范围必须在放款日之后为活跃贷款
            if self.p[active_status] == active:
                #贷款结束日和期间最后一天做比较，取最早的一个
                if self.date_sub(self.p[end], self.p[putout_end]) >= 0:
                    self.p[end] = self.p[putout_end]
                if self.date_sub(self.p[last_end], self.p[putout_end]) >= 0:
                    self.p[last_end] = self.p[putout_end]
                #贷款状态修正判定, 获得计息截止日
                self.p[loan_status], self.p[actual_end_date] = self.get_loan_status()
                #修正最早的日期作为计息截止日
                self.get_real_actual_end_date()
                self.report[period][loan_status], self.report[period][actual_end_date] = self.p[loan_status], self.p[actual_end_date]

                #获得以前年度计息截止日
                self.report[period][pre_actual_end_date] = self.p[pre_actual_end_date] = self.get_pre_actual_date()
                #获得时间点的科目信息数值
                account_list = [premium_receivable_principal, contract_assets_early_repay]
                for account in account_list:
                    self.report[period][account] = self.get_time_point_value(self.p[end], account)
                    self.report[period + '_before'][account] = self.get_time_point_value(self.p[last_end], account)

                #担保收入 - 担保累加
                self.report[period][unearned_premium_reserve] = self.subtract_value(self.p[total_guarantee],
                                                                self.get_accumulate_value(BEGIN, self.p[end], guarantee))
                #以前期间担保收入 - 担保累加
                self.report[period + '_before'][unearned_premium_reserve] = self.subtract_value(self.p[total_guarantee],
                                                                            self.get_accumulate_value(BEGIN, self.p[last_end], guarantee))
                #获取累加的科目信息数值
                account_list = [cash_for_matching, post_origination, cash_for_post_origination,
                                cash_for_guarantee, match_early_repay, revenue_margin]
                #计算当前期间数据
                value = self.get_accumulate_value_list(self.p[start], self.p[actual_end_date], account_list)
                #计算以前期间数据
                pre_value = self.get_accumulate_value_list(BEGIN, self.p[pre_actual_end_date], account_list)
                for account in account_list:
                    self.report[period][account] = value[account]
                    self.report[period + '_before'][account] = pre_value[account]

                #获取累加的科目信息数值 - 使用贷款终止日作为计算
                account_list = [guarantee, interest_adjust_for_guarantee, interest_adjust_for_match_early_repay]
                #计算当前期间数据
                value = self.get_accumulate_value_list(self.p[start], self.p[end], account_list)
                #计算以前期间数据
                pre_value = self.get_accumulate_value_list(BEGIN, self.p[last_end], account_list)
                for account in account_list:
                    self.report[period][account] = value[account]
                    self.report[period + '_before'][account] = pre_value[account]

                #判断该分录是否需要结转
                self.report[period][carry] = self.p[carry] = self.get_carry_value()

                #报表分录调整前，调整，调整后初始化
                self.report[period][unadjusted] = {}
                self.report[period][adjust_num] = {}
                self.report[period][adjusted] = {}

                '''
                    调整前处理
                    借正贷负操作     
                '''
                #当前期间借方处理
                for account in [premium_receivable_principal, revenue_margin, cash_for_guarantee,
                                cash_for_post_origination, contract_assets_early_repay, cash_for_matching]:
                    self.report[period][unadjusted][account] = self.report[period][account]
                #当前期间贷方处理
                for account in [guarantee, unearned_premium_reserve, post_origination, match_early_repay,
                                interest_adjust_for_match_early_repay, interest_adjust_for_guarantee]:
                    self.report[period][unadjusted][account] = -self.report[period][account]
                self.report[period][unadjusted][guarantee_liability_460] = self.report[period][unadjusted][unearned_premium_reserve]

                #以前期间借方处理
                for account in [cash_for_guarantee, cash_for_post_origination, cash_for_matching]:
                    self.report[period][unadjusted]['pre_'+account] = self.report[period + '_before'][account]

                #以前期间贷方处理
                for account in [guarantee, interest_adjust_for_guarantee, post_origination,
                                match_early_repay, interest_adjust_for_match_early_repay]:
                    self.report[period][unadjusted]['pre_'+account] = -self.report[period + '_before'][account]
                self.report[period][unadjusted][guarantee_loss] = self.multiply_value(self.p[loss_ratio] / self.p[guarantee_ratio],
                                        (-self.report[period][unadjusted][guarantee]))
                self.report[period][unadjusted][pre_guarantee_loss] = self.multiply_value(self.p[loss_ratio] / self.p[guarantee_ratio],
                                                                                         (-self.report[period][unadjusted][pre_guarantee]))

                #对于太保模式的机构，进行担保相关科目调整前和调整数的归零调整
                if self.get_is_include(self.p[lineid], config['FILTER1']):

                    for account in [premium_receivable_principal, cash_for_guarantee, guarantee,
                                    interest_adjust_for_guarantee, guarantee_liability_460, impairment_loss1,
                                    pre_impairment_loss1, pre_cash_for_guarantee,
                                    pre_guarantee, pre_interest_adjust_for_guarantee]:
                        if account in [premium_receivable_principal, cash_for_guarantee, guarantee_liability_460,
                                       pre_cash_for_guarantee, pre_guarantee]:
                            try:
                                self.report[period][unadjusted][account] = self.report[period][unadjusted][account] * 0
                                self.report[period][adjust_num][account] = self.report[period][adjust_num][account] * 0
                            except:
                                self.report[period][unadjusted][account] = 0
                                self.report[period][adjust_num][account] = 0
                        elif account in [interest_adjust_for_guarantee, pre_interest_adjust_for_guarantee]:
                            self.report[period][unadjusted][account] = self.report[period][unadjusted][account] * 0.65
                        else:
                            self.report[period][unadjusted][account] = 0
                            self.report[period][adjust_num][account] = 0

                #调整后数据初始化
                total_account = [premium_receivable_principal, cash_for_guarantee, guarantee,
                                 interest_adjust_for_guarantee,
                                 guarantee_liability_460, pre_cash_for_guarantee, pre_guarantee,
                                 pre_interest_adjust_for_guarantee, cash_for_post_origination, post_origination,
                                 pre_cash_for_post_origination, pre_post_origination, contract_assets_early_repay,
                                 cash_for_matching, match_early_repay, interest_adjust_for_match_early_repay,
                                 pre_cash_for_matching, pre_match_early_repay,
                                 pre_interest_adjust_for_match_early_repay,
                                 revenue_margin, guarantee_loss, pre_guarantee_loss]
                for account in total_account:
                    self.report[period][adjusted][account] = self.report[period][unadjusted][account]

                #未来损失率计算，调整后数据初始化
                total_account = [future_loss_month, pre_future_loss_month, future_loss_ratio, pre_future_loss_ratio,
                                 premium_receivable_day1, contract_assets_day1,
                                 # premium_receivable_loss, contract_assets_loss,pre_premium_receivable_loss, pre_contract_assets_loss,
                                 ]
                for account in total_account:
                    self.report[period][adjusted][account] = self.report[period][account]

                '''对调整后 收入报表 进行 拨备计算'''
                #判断是不是放款月
                if self.get_include_month(self.p[putout_date], self.p[start], self.p[end]):
                    self.report[period][adjusted][fair_value_guarantees] = self.p[total_guarantee]
                    #放款月实际天数
                    self.p[actual_day_num] = self.date_sub(self.p[end], self.p[putout_date])
                    #之前期间的累积天数
                    self.p[accumulate_day_num] = 0
                else:
                    self.report[period][adjusted][pre_fair_value_guarantees] = self.p[total_guarantee]
                    #贷款月实际天数
                    if self.date_sub(self.p[putout_end], self.p[start]) < 0:
                        self.p[actual_day_num] = 0
                    else:
                        self.p[actual_day_num] = self.date_sub(self.p[end], self.p[start]) + 1

                    #之前期间的累积天数
                    if self.date_sub(self.p[putout_end], self.p[start]) < 0:
                        self.p[accumulate_day_num] = self.date_sub(self.p[putout_end], self.p[putout_date])
                    else:
                        self.p[accumulate_day_num] = self.date_sub(self.p[start], self.p[putout_date]) - 1

                #判断是不是计划结束月
                if self.get_include_month(self.p[putout_end], self.p[start], self.p[end]):
                    #结束月实际天数
                    self.p[actual_day_num] = self.date_sub(self.p[putout_end], self.p[start]) + 1

                #计算计入利润表拨备
                self.report[period][adjusted][charge_income] = self.multiply_value(float(self.p[principal] * self.p[loss_ratio] / self.p[day_num]),
                                                                                   self.p[actual_day_num])
                #计算以前期间计入利润表拨备
                self.report[period][adjusted][pre_charge_income] = self.multiply_value(float(self.p[principal] * self.p[loss_ratio] / self.p[day_num]),
                                                                                       self.p[accumulate_day_num])

                #更改特别合作机构的输出值
                if self.get_is_include(self.p[lineid], config['FILTER1']):
                    # 调整后的字段归零
                    for account in [fair_value_guarantees, pre_fair_value_guarantees, charge_income, pre_charge_income]:
                        self.report[period][adjusted][account] = 0

                    # 调整后的字段乘以比例 TPY_Model_ratio2
                    for st in [cash_for_post_origination, post_origination, pre_cash_for_post_origination, pre_post_origination,
                               contract_assets_early_repay, cash_for_matching, match_early_repay, pre_cash_for_matching,
                               pre_match_early_repay, charge_income, pre_fair_value_guarantees, pre_charge_income]:
                        self.report[period][adjusted][st] = self.report[period][adjusted][st] * self.p[TPY_Model_ratio2]

                    # 乘以比例 TPY_Model_ratio1
                    self.report[period][reduction_of_interest_matching_accumulate] = self.report[period][reduction_of_interest_matching_accumulate] * self.p[TPY_Model_ratio1]
                    self.report[period][reduction_of_interest_matching_current] = self.report[period][reduction_of_interest_matching_current] * self.p[TPY_Model_ratio1]

                    # 调整后的字段 乘以比例 TPY_Model_ratio1
                    for st in [interest_adjust_for_match_early_repay, pre_interest_adjust_for_match_early_repay]:
                        self.report[period][adjusted][st] = self.report[period][adjusted][st] * self.p[TPY_Model_ratio1]

                # 大地模式输出值修改
                if self.get_is_include(self.p[lineid], config['FILTER2']):
                    if fair_value_guarantees not in self.report[period][adjusted].keys():
                        self.report[period][adjusted][fair_value_guarantees] = 0
                    if pre_fair_value_guarantees not in self.report[period][adjusted].keys():
                        self.report[period][adjusted][pre_fair_value_guarantees] = 0

                    # 需要修改的调整后字段
                    for account in [premium_receivable_principal, cash_for_guarantee, guarantee,
                                    interest_adjust_for_guarantee, guarantee_liability_460, # impairment_loss1,pre_impairment_loss1,
                                    pre_cash_for_guarantee, pre_guarantee, pre_interest_adjust_for_guarantee,
                                    fair_value_guarantees, pre_fair_value_guarantees, charge_income, pre_charge_income,
                                    premium_receivable_day1]:
                        # 输出值 乘以 0.5
                        self.report[period][adjusted][account] = self.report[period][adjusted][account] * 0.5

                    # 需要修改的调整后字段
                    for st in [cash_for_post_origination, post_origination, pre_cash_for_post_origination, pre_post_origination,
                               contract_assets_early_repay, cash_for_matching, match_early_repay, pre_cash_for_matching,
                               pre_match_early_repay, contract_assets_day1]:
                        # 输出值 乘以比例 TPY_Model_ratio4
                        self.report[period][adjusted][st] = self.report[period][adjusted][st] * self.p[TPY_Model_ratio4]

                    # 需要修改的字段
                    for st in [reduction_of_interest_guarantee_accumulate, reduction_of_interest_guarantee_current]:
                        # 输出值 乘以 0.5
                        self.report[period][st] = self.report[period][st] * 0.5

                    # 输出值 乘以比例 TPY_Model_ratio3
                    self.report[period][reduction_of_interest_matching_accumulate] = self.report[period][reduction_of_interest_matching_accumulate] * self.p[TPY_Model_ratio3]
                    self.report[period][reduction_of_interest_matching_current] = self.report[period][reduction_of_interest_matching_current] * self.p[TPY_Model_ratio3]

                    # 调整后的字段 输出值乘以比例 TPY_Model_ratio3
                    for st in [interest_adjust_for_match_early_repay, pre_interest_adjust_for_match_early_repay]:
                        self.report[period][adjusted][st] = self.report[period][adjusted][st] * self.p[TPY_Model_ratio3]

    '''
        定义输出字段，可以对需要输出的字段进行调整修改
            h1 为参数字段
            h2 为收入报表计算输出字段，包含调整前，调整，调整后
            h3 记录创建人和时间
    '''
    def output_report_list(self, writer):
        # h1 h2 h3 收入报表表头
        h1 = [putout_no, putout_date, loan_term, business_sum, itemname, ctype,
              guarantee_ratio, margin_ratio, loss_ratio,
              overdue_date, overdue_days, mob, finish_date, buy_date, date, adjust_status]

        h2 = [loan_status, premium_receivable_principal, cash_for_guarantee, guarantee, interest_adjust_for_guarantee,
              guarantee_liability_460,
              pre_cash_for_guarantee, pre_guarantee, pre_interest_adjust_for_guarantee, cash_for_post_origination,
              post_origination, pre_cash_for_post_origination, pre_post_origination, contract_assets_early_repay,
              cash_for_matching, match_early_repay, interest_adjust_for_match_early_repay, pre_cash_for_matching,
              pre_match_early_repay,
              pre_interest_adjust_for_match_early_repay,
              fair_value_guarantees, pre_fair_value_guarantees, charge_income, pre_charge_income,  # 拨备字段
              # future_loss_month, pre_future_loss_month, future_loss_ratio, pre_future_loss_ratio,
              premium_receivable_day1, contract_assets_day1,
              # premium_receivable_loss, contract_assets_loss, pre_premium_receivable_loss, pre_contract_assets_loss,
              month_of_loan_end, month_of_loan_start,
              implied_price_concession_ratio_end, implied_price_concession_ratio_start,
              reduction_of_interest_guarantee_accumulate, reduction_of_interest_guarantee_current,
              reduction_of_interest_matching_accumulate, reduction_of_interest_matching_current,
              start, end, actual_end_date]

        h3 = [created_by, created_date]

        #合并表头字段
        if config['HEADER'] == 1:
            head = []
            head.extend(h1)
            head.extend(h2)
            head.extend(h3)
            writer.writerow(head)

        #循环输出每个期间的收入数据
        for period in self.period_list.keys():
            table = []
            if self.report[period][active_status] == active:
                #调整前，调整，调整后
                status = [unadjusted, adjust_num, adjusted]
                for st in status:
                    line = []
                    if st in [unadjusted, adjust_num]:
                        pass
                    else:
                        for key1 in h1:
                            #循环h1字段
                            if key1 == date:
                                line.append(period)
                            elif key1 == adjust_status:
                                line.append(st)
                            else:
                                line.append(self.p[key1])

                        for key2 in h2:
                            #循环h2字段
                            if key2 in self.report[period][st].keys():
                                line.append(self.report[period][st][key2])
                            elif key2 in [start, end, actual_end_date, loan_status]:

                                line.append(self.report[period][key2])
                            elif key2 in [month_of_loan_end, month_of_loan_start,
                                          implied_price_concession_ratio_end, implied_price_concession_ratio_start,
                                          reduction_of_interest_guarantee_accumulate,
                                          reduction_of_interest_guarantee_current,
                                          reduction_of_interest_matching_accumulate,
                                          reduction_of_interest_matching_current]:
                                line.append(self.report[period][key2])
                            else:
                                line.append(0)
                        for key3 in h3:
                            #循环h3字段
                            if key3 == created_by:
                                line.append(PROGRAM)
                            elif key3 == created_date:
                                line.append(TIME)
                        table.append(line)
                #表格数据写入文本
                writer.writerows(table)

    '''模型计算主程序'''
    def calculation_program(self, schedule_writer=None, report_writer=None):
        #第一层：现金流计算程序
        self.cash_flow_calculation()
        #第二层：收入计算程序
        self.revenue_calculation()
        #第三层：资产负债表计算程序
        self.balance_sheet_calculation()

        #导出还款计划表
        if config['OUTPUT_SCHEDULE'] == 1:
            self.output_schedule_list(schedule_writer)

        #个别合作机构不做第二步运算
        if not self.get_is_include(self.p[lineid], config['CANCEL']):

            if config['OUTPUT_REPORT'] == 1:
                #收入报表计算
                self.report_calculation()

            if config['OUTPUT_REPORT'] == 1:
                #导出收入和拨备报表
                self.output_report_list(report_writer)

            if config['HEADER'] == 1:
                """定义全局变量"""
                # global config
                """每个文件表头字段只输出一次"""
                config['HEADER'] = 0
################################################数据执行#################################################################

class DataEngine(object):
    """导入，执行操作，导出数据"""

    def __init__(self):

        """数据导入"""
        path = os.getcwd() + os.sep + 'basic_data'
        if os.path.exists(path+'.xlsx'):
            self.type = 'Excel'
        elif os.path.exists(path+'.txt'):
            self.type = 'Text'
        else:
            print('Error: No such file basic_data!')
            exit(1)
        """初始化参数列表"""

        # 通过静态方法给 类变量 future_loss 赋值
        self.future_loss = self.get_future_loss()
        # 通过静态方法给 类变量 loss_ratio_margin 赋值
        self.loss_ratio_margin = self.get_lossratio_and_margin()
        # 初始化数据列表
        self.param_list = []

    # 从config.xls配置第二页中读取future_loss
    @staticmethod
    def get_future_loss():
        configPath = os.getcwd() + os.sep + 'config.xls'
        workbook = xlrd.open_workbook(configPath)
        worksheet = workbook.sheet_by_index(1)
        rows = worksheet.nrows      # 获取工作表的行数
        cols = worksheet.ncols      # 获取工作表的列数

        future_loss = {}    # 初始化 future_loss 字典
        # 遍历工作表的行数
        for r in range(1, rows):
            quater = worksheet.cell_value(r, 2)     # 获得季度
            future_loss[quater] = {}
            # 季度开始日期
            future_loss[quater][start] = xlrd.xldate_as_datetime(worksheet.cell_value(r, 0), 0).strftime("%Y/%m/%d")
            # 季度结束日期
            future_loss[quater][end] = xlrd.xldate_as_datetime(worksheet.cell_value(r, 1), 0).strftime("%Y/%m/%d")
            for c in range(3, cols):
                month = c - 2
                # 1-36个月的 future_loss 值
                future_loss[quater][month] = worksheet.cell_value(r, c)
        return future_loss

    # 从config.xls配置第三页中读取 Implied price concession ratio
    @staticmethod
    def get_IPCR():
        path = os.getcwd() + os.sep + 'config.xls'
        excel = xlrd.open_workbook(path)
        sheet = excel.sheet_by_index(2)
        rows = sheet.nrows
        ipcr = {}

        for row in range(0, rows):
            ipcr[row] = sheet.cell_value(row, 1)
        return ipcr

    # 从config.xls配置第四页中读取 annual_loss_ratio 和 annual_margin
    @staticmethod
    def get_lossratio_and_margin():
        configPath = os.getcwd() + os.sep + 'config.xls'
        excel = xlrd.open_workbook(configPath)
        sheet = excel.sheet_by_index(3)
        rows = sheet.nrows
        dic = {}

        for row in range(1, rows):
            dic1 = {'annual_loss_ratio': sheet.cell_value(row, 3)}      # 临时字典1存放 annual_loss_ratio
            dic2 = {'annual_margin': sheet.cell_value(row, 4)}          # 临时字典2存放 annual_margin
            dic.setdefault(sheet.cell_value(row, 2), [dic1, dic2])      # 每个季度作为键与相应的列表[dic1, dic2]作为值构建字典
        return dic

    '''获取数据源，根据不同的类型进行数据接入'''
    def get_data(self):
        if self.type == "Excel":
            #读取excel的方式
            path = os.getcwd() + os.sep + 'basic_data.xlsx'
            excel = xlrd.open_workbook(path)
            sheet = excel.sheet_by_index(0)
            row, col = sheet.nrows, sheet.ncols     # 获得 basic_data 数据表的行数和列数

            # 遍历行数，对每一行的数据处理
            for r in range(1, row):
                line = []   # 初始化临时变量，用来保存一行数据
                # 遍历一行数据的列数
                for c in range(0, col):
                    if c == 5:  # 数据的第6列为日期格式，若是第6列需要转换以下格式，然后放到临时变量line中
                        line.append(xlrd.xldate_as_datetime(sheet.cell_value(r, c), 0))
                    else:
                        # 如果不是第6列，则直接放到临时变量line中
                        line.append(sheet.cell_value(r, c))
                # 内循环结束，line已经存放了一行的数据，将line添加到类变量 self.param_list 中
                self.param_list.append(line)

        elif self.type == "Text":
            #读取文本方式
            path = os.getcwd() + os.sep + 'basic_data.txt'
            file = op(path, 'r', 'gbk')
            rowsText = file.readlines()     # readlines读取全部行的内容

            # 逐行处理
            for row in rowsText:
                # 对每一行使用制表符进行切片，形成列表格式
                data = row.split("\t")
                # 如果不是第一行，就添加到类变量 self.param_list 中
                if data[0] not in ['putout_no', '出账编号']:
                    self.param_list.append(data)


    '''获取业务参数修改历史'''
    @staticmethod
    def get_changelog():
        # 读取config.xls配置表的第一页和第五页
        configPath = os.getcwd() + os.sep + 'config.xls'
        excel = xlrd.open_workbook(configPath)
        busConfig = excel.sheet_by_index(0)     # 第一页业务参数配置
        changeLog = excel.sheet_by_index(4)     # 第五页参数修改历史
        count = changeLog.nrows     # 获得第五页的行数

        '''
        参数字典，以列表方式存储
            upfront_cost
            ongoing_cost
            match_early_rate
            TPY_model_ratio1
            TPY_model_ratio2
            TPY_model_ratio3
            TPY_model_ratio4
        '''
        dicList = [{}, {}, {}, {}, {}, {}, {}]      # 初始化业务参数列表

        # 先对第五页做处理，遍历行数
        for i in range(1, count):
            # 对第一列的日期格式做转换处理
            time = xlrd.xldate_as_datetime(changeLog.cell_value(i, 0), 0).strftime('%Y/%m/%d')
            time = datetime.datetime.strptime(time, '%Y/%m/%d')
            # 遍历 2-8 列（分别为7个业务参数）
            for j in range(1, 8):
                try:
                    if changeLog.cell_value(i, j) != '':
                        # 若是参数有修改值，则以（日期：修改值）作为键值对存到参数列表
                        dicList[j-1][time] = changeLog.cell_value(i, j)
                except:
                    pass

        # 初始时间
        ever = '2015/01/01'
        ever = datetime.datetime.strptime(ever, '%Y/%m/%d')
        t = 0
        # 对第一页做处理
        for r in range(0, 7):
            # 以（初始时间：业务参数值）作为键值对存到参数列表
            dicList[t][ever] = busConfig.cell_value(r, 1)
            t += 1

        return dicList

    #统一日期格式 xxxx/xx/xx
    @staticmethod
    def date_format(param_date):
        if param_date == '' or param_date is None:
            param_date = ''
        elif len(param_date) > 10:
            param_date = param_date[:10]
            param_date = param_date.replace('-', '/')
        else:
            splits = ['-', '/']
            for split in splits:
                if split in param_date:
                    try:
                        param_date = datetime.datetime.strptime(param_date, '%Y'+split+'%m'+split+'%d')
                        param_date = datetime.datetime.strftime(param_date, '%Y/%m/%d')
                    except:
                        print('Param error: datetime format error!')
                        exit(1)
        return param_date

    #判断putout_date属于哪个季度
    @staticmethod
    def get_sornum(putout_date):
        putoutDate = putout_date.split('/')     # 对传入的日期做切片处理，生成日期列表，格式为['xxxx', 'xx', 'xx']

        if int(putoutDate[1]) <= 3:             # 若列表的第二个值即月份小于等于3，则为第一个季度
            sornum = putoutDate[0] + 'Q1'       # 季度 = 年份 + 'Q1'

        elif int(putoutDate[1]) <= 6:
            sornum = putoutDate[0] + 'Q2'       # 季度 = 年份 + 'Q2'

        elif int(putoutDate[1]) <= 9:
            sornum = putoutDate[0] + 'Q3'       # 季度 = 年份 + 'Q3'

        else:
            sornum = putoutDate[0] + 'Q4'       # 季度 = 年份 + 'Q4'
        return sornum

    # 计算 guarantee_ratio 和 fulltime_margin
    def get_guarantee_and_margin(self, sornum, loan_term):
        # 传入参数为季度和借款期限
        loss_and_margin = self.loss_ratio_margin[sornum]    #根据相应的季度从类变量 self.loss_ratio_margin 中得到 annual_loss 和 anual_margin 的值
        loss, margin = loss_and_margin[0], loss_and_margin[1]   # 将值分别赋给 loss 和 margin

        # fulltime_margin = annual_margin * 1.67 * (借款期限 / 36)
        fulltime_margin = margin['annual_margin'] * 1.67 * (loan_term / 36)

        # 普通机构 guarantee_ratio = annual_loss_ratio * 1.67 * (借款期限 / 36) * 0.93 + fulltime_margin
        guarantee_ratio = loss['annual_loss_ratio'] * 1.67 * (loan_term / 36) * 0.93 + fulltime_margin
        # 太保模式 guarantee_ratio_taibao = annual_loss_ratio * 1.67 * 1.04 * 0.93
        guarantee_ratio_taibao = loss['annual_loss_ratio'] * 1.67 * 1.04 * 0.93

        return guarantee_ratio, guarantee_ratio_taibao, fulltime_margin

    '''根据进程序号分配参数列表'''
    def assign_param_list(self, num):
        assign_param = {}   # 初始化进程参数字典
        n = 1
        # 构建进程参数字典，其格式为 assign_param = {1: [参数列表], 2: [参数列表], 3: [参数列表], 4: [参数列表]}
        for i in range(1, num + 1):
            assign_param[i] = []
        # 遍历类变量 self.param_list 参数列表
        for param in self.param_list:
            # 将每一条数据分配给 assign_param[进程号][参数列表]
            assign_param[n].append(param)
            n = n + 1
            if n == num + 1:
                n = 1
        print("********************************************************")
        print('Task[' + str(num) + "] ************************** get assign param list")
        print("********************************************************")
        return assign_param

    '''装入参数'''
    def input_param(self, param_list):
        # 赋值给类变量
        self.param_list = param_list

    '''配置业务参数'''
    @staticmethod
    def readConfig():
        # 全局变量 config 仅包括了系统配置，需要将业务参数配置添加到config中
        configPath = os.getcwd() + os.sep + 'config.xls'
        if not os.path.exists(configPath):
            print('File config.xls does not exist!')
            exit(1)
        excel = xlrd.open_workbook(configPath)
        sheet = excel.sheet_by_index(0)
        rows = sheet.nrows  # 获得第一页的行数

        # 遍历行数
        for row in range(0, rows):
            # 若为 FILTER1，FILTER2， CANCEL 机构代码，则做如下处理
            if sheet.cell_value(row, 0) == 'FILTER1' or sheet.cell_value(row, 0) == 'CANCEL' or sheet.cell_value(row, 0) == 'FILTER2':
                li = sheet.cell_value(row, 1).split(',')    # 对机构代码以','切分，形成列表 ['机构代码', '机构代码', '机构代码']
                config[sheet.cell_value(row, 0)] = li       # 放到全局变量字典中

            # 若为 SERVER 机构代码，则做如下入理
            elif sheet.cell_value(row, 0) == 'SERVER':
                listd = sheet.cell_value(row, 1).split(',')     # 以','切分
                dic = {}    # 临时变量
                # 遍历 listd 列表
                for ld in listd:
                    # 对 listd 列表中的每一项都转换成字典形式，放到 dic 中
                    ld = ld.split(':')
                    dic[ld[0]] = int(ld[1])
                config[sheet.cell_value(row, 0)] = dic
            else:
                config[sheet.cell_value(row, 0)] = sheet.cell_value(row, 1)

    '''开始创建数据'''
    def create_data(self, process_num, period=None):

        '''生成报表文件'''
        if config['OUTPUT_SCHEDULE'] == 1:
            schedule_name = config['OUTPUT_SCHEDULE_FILE'] + '_' + str(process_num) + ".txt"
            schedule_file = open(schedule_name, 'w')
            schedule_writer = csv.writer(schedule_file, delimiter='\t', lineterminator='\r')
        else:
            schedule_writer = None

        if config['OUTPUT_REPORT'] == 1:
            report_name = config['OUTPUT_REPORT_FILE'] + '_' + str(process_num) + ".txt"
            report_file = open(report_name, 'w')
            report_writer = csv.writer(report_file, delimiter='\t', lineterminator='\r')
        else:
            report_writer = None

        #获取未来损失率和参数修改历史
        future_loss = self.future_loss
        changeList = self.get_changelog()
        ipcr = self.get_IPCR()
        #读取配置参数
        self.readConfig()

        n = 0
        param_num = len(self.param_list)

        '''循环每一笔贷款的参数'''
        for param in self.param_list:
            if param[7] == '':
                continue
            #清洗字段，统一日期格式 %Y/%m/%d
            if len(param) == 16:
                del param[11:13]
            param[-1] = param[-1].replace('\r\n', '')
            for i in [5, 11, 12, 13]:
                param[i] = self.date_format(str(param[i]))

            #损失率计算
            sornum = self.get_sornum(param[5])
            param[3] = float(param[3])

            tuple_ratio = self.get_guarantee_and_margin(sornum, int(param[3]))      # 获得(guarantee_ratio, guarantee_ratio_taibao, fulltime_margin)
            # 将这三个参数放到 param 末尾
            param.append(tuple_ratio[0])
            param.append(tuple_ratio[1])
            param.append(tuple_ratio[2])

            #对每一笔贷款进行计算
            revenue = Revenue(config, param, period, future_loss, changeList, ipcr)     # 传入参数，实例化类 Revenue
            revenue.calculation_program(schedule_writer, report_writer)                 # 调用类方法，报表计算主函数

            #程序运行时间相关
            n += 1
            print('Task[{0}] Runing：{1}'.format(process_num, param))
            secends = time.process_time()  # 循环一次程序运行时间
            percent = str(round((n / param_num) * 100, 2)) + '%'  # 已经计算结束占全部的百分比
            times = str(int(secends / 60)) + ' minutes ' + str(int(secends % 60)) + ' secends'  # 已经花费的时间
            avg_cost = round(secends / n, 1)    # 平均花费的时间
            ex_cost = round(avg_cost * (param_num-n), 1)
            ex_time = str(int(ex_cost / 60)) + ' minutes ' + str(int(ex_cost % 60)) + ' secends'    # 预计还要花费的时间
            print('Percent：{0} | Usedtime：{1} | Will take：{2}'.format(percent, times, ex_time))

################################################主程序区#################################################################

'''设置计算期间数据'''
def set_period(dt_list):
    period_list = {}
    #"2017-01-2017-06"
    for dt in dt_list:
        period_list[dt] = {}
        num = len(dt.split("-"))
        if num == 2:
            period_list[dt][start] = dt[0:4] + '/' + dt[5:7] + '/01'
            period_list[dt][end] = dt[0:4] + '/' + dt[5:7] + '/' + str(monthrange(int(dt[0:4]), int(dt[5:7]))[1])
        elif num == 4:
            period_list[dt][start] = dt[0:4] + '/' + dt[5:7] + '/01'
            period_list[dt][end] = dt[8:12] + '/' + dt[13:15] + '/' + str(monthrange(int(dt[8:12]), int(dt[13:15]))[1])
        else:
            period_list[dt][start] = dt[0:4] + '/01/01'
            period_list[dt][end] = dt[0:4] + '/12/31'
        """返回时间段列表"""
    print("准备计算的期间数为：")
    print(period_list)
    return period_list

'''获取文件夹路径'''
def combine_file(name, num):
    file = open(name + '.txt', 'w')
    writer = csv.writer(file, delimiter='\t')
    """遍历文件名"""
    for i in range(1, num+1):
        schdule_path = name + '_' + str(i) + '.txt'
        for line in open(schdule_path, 'r'):
            line = line.replace("\n", "").split('\t')
            writer.writerow(line)
        os.remove(schdule_path)
    file.close()
    print('The ' + name + ' file finished')

'''单任务处理'''
def single_task(period, num=1):
    data = DataEngine()
    data.get_data()
    data.create_data(1, period)
    if config['OUTPUT_SCHEDULE'] == 1:
        #导出还款计划表
        if config['COMBINE'] == 1:
            combine_file(config['OUTPUT_SCHEDULE_FILE'], num)
    if config['OUTPUT_REPORT'] == 1:
        #导出收入报表
        if config['COMBINE'] == 1:
            combine_file(config['OUTPUT_REPORT_FILE'], num)

def main_program(num,param_list=None, period=None):
    #主程序
    data = DataEngine()
    data.input_param(param_list)
    data.create_data(num, period)

'''多任务处理'''
def multi_task(period, num):

    data = DataEngine()
    data.get_data()
    #根据总参数数量平均分配
    assgin_param = data.assign_param_list(num)

    multiprocessing.freeze_support()
    print('Parent process %s.' % os.getpid())
    p = multiprocessing.Pool()

    #多个进程循环启动
    for i in range(1, num + 1):
        param_list = assgin_param[i]
        p.apply_async(main_program, args=(i, param_list, period))

    print('Waiting for all subprocesses done...')
    p.close()
    p.join()
    print('All subprocesses done.')

    if config['OUTPUT_SCHEDULE'] == 1:
        #导出还款计划表
        if config['COMBINE'] == 1:
            combine_file(config['OUTPUT_SCHEDULE_FILE'], num)
    if config['OUTPUT_REPORT'] == 1:
        #导出收入报表
        if config['COMBINE'] == 1:
            combine_file(config['OUTPUT_REPORT_FILE'], num)

    print('finished work!')

if __name__ == '__main__':
    #设置计算日期
    PERIOD = set_period(['2014-12-2018-12'])
    if config['PROCESS'] == 1:
        single_task(PERIOD, 1)
    else:
        multi_task(PERIOD, config['PROCESS'])
