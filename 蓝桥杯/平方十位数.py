'''
    题目描述
    由0~9这10个数字不重复、不遗漏，可以组成很多10位数字。
    这其中也有很多恰好是平方数（是某个数的平方）。
    
    比如：1026753849，就是其中最小的一个平方数。
    
    请你找出其中最大的一个平方数是多少？
    
    注意：你需要提交的是一个10位数字，不要填写任何多余内容。
'''
def tenByteNum():
    maxNum = 0
    t = 100000
    while True:
        if t*t < 9876543210:
            maxNum = t
            break
        t -= 1

    while True:
        powNum = pow(maxNum, 2)
        if len(set(str(powNum))) == 10:
            print('Num：', maxNum)
            print('最大的不重复平方数：', powNum)
            break
        maxNum -= 1

if __name__ == '__main__':
    tenByteNum()
