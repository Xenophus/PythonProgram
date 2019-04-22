import numpy as np

'''
    题目描述
    康威生命游戏是英国数学家约翰·何顿·康威在1970年发明的细胞自动机。  
    这个游戏在一个无限大的2D网格上进行。
    
    初始时，每个小方格中居住着一个活着或死了的细胞。
    下一时刻每个细胞的状态都由它周围八个格子的细胞状态决定。
    
    具体来说：
    
    1. 当前细胞为存活状态时，当周围低于2个（不包含2个）存活细胞时， 该细胞变成死亡状态。（模拟生命数量稀少）
    2. 当前细胞为存活状态时，当周围有2个或3个存活细胞时， 该细胞保持原样。
    3. 当前细胞为存活状态时，当周围有3个以上的存活细胞时，该细胞变成死亡状态。（模拟生命数量过多）
    4. 当前细胞为死亡状态时，当周围有3个存活细胞时，该细胞变成存活状态。 （模拟繁殖）
    
    当前代所有细胞同时被以上规则处理后, 可以得到下一代细胞图。按规则继续处理这一代的细胞图，可以得到再下一代的细胞图，周而复始。
    
    例如假设初始是:(X代表活细胞，.代表死细胞)
    .....
    .....
    .XXX.
    .....
    
    下一代会变为:
    .....
    ..X..
    ..X..
    ..X..
    .....
    
    康威生命游戏中会出现一些有趣的模式。例如稳定不变的模式：
    
    ....
    .XX.
    .XX.
    ....
    
    还有会循环的模式：
    
    ......      ......       ......
    .XX...      .XX...       .XX...
    .XX...      .X....       .XX...
    ...XX.   -> ....X.  ->   ...XX.
    ...XX.      ...XX.       ...XX.
    ......      ......       ......
    
    
    本题中我们要讨论的是一个非常特殊的模式，被称作"Gosper glider gun"：
    
    ......................................
    .........................X............
    .......................X.X............
    .............XX......XX............XX.
    ............X...X....XX............XX.
    .XX........X.....X...XX...............
    .XX........X...X.XX....X.X............
    ...........X.....X.......X............
    ............X...X.....................
    .............XX.......................
    ......................................
    
    假设以上初始状态是第0代，请问第1000000000(十亿)代一共有多少活着的细胞？
    
    注意：我们假定细胞机在无限的2D网格上推演，并非只有题目中画出的那点空间。
    当然，对于遥远的位置，其初始状态一概为死细胞。
    
    注意：需要提交的是一个整数，不要填写多余内容。
'''
def lifeGame(times):
    arr = np.zeros([11, 38])
    arr = arr.astype(int)
    arr[1, 25] = 1
    arr[2, 23:25] = 1
    arr[3, [13, 14, 21, 22, 35, 36]] = 1
    arr[4, [12, 16, 21, 22, 35, 36]] = 1
    arr[5, [1, 2, 11, 17, 21, 22]] = 1
    arr[6, [1, 2, 11, 15, 17, 18, 23, 24]] = 1
    arr[7, [11, 17, 25]] = 1
    arr[8, [12, 16]] = 1
    arr[9, 13:15] = 1
    print('原始生命：\n', arr)

    time = times
    while(times):
        times -= 1
        a = arr.copy()
        for row in range(0, 5):
            for col in range(0, 5):
                if row == 0:
                    if col == 0:
                        lifeNum = a[row, col+1] + a[row+1, col] + a[row+1, col+1]
                    elif col == 4:
                        lifeNum = a[row, col-1] + a[row+1, col-1] + a[row+1, col]
                    else:
                        lifeNum = a[row, col-1] + a[row, col+1] + a[row+1, col-1] + a[row+1, col] + a[row+1, col+1]
                elif row == 4:
                    if col == 0:
                        lifeNum = a[row-1, col] + a[row-1, col+1] + a[row, col+1]
                    elif col ==4:
                        lifeNum = a[row-1, col-1] + a[row-1, col] + a[row, col-1]
                    else:
                        lifeNum = a[row-1, col-1] + a[row-1, col] + a[row-1, col+1] + a[row, col-1] + a[row, col+1]
                else:
                    if col == 0:
                        lifeNum = a[row-1, col] + a[row-1, col+1] + a[row, col+1] + a[row+1, col] + a[row+1, col+1]
                    elif col == 4:
                        lifeNum = a[row-1, col-1] + a[row-1, col] + a[row, col-1] + a[row+1, col-1] + a[row+1, col]
                    else:
                        lifeNum = a[row-1, col-1] + a[row-1, col] + a[row-1, col+1] + \
                                  a[row, col-1] + a[row, col+1] + \
                                  a[row+1, col-1] + a[row+1, col] + a[row+1, col+1]

                if a[row, col] == 1:
                    if lifeNum < 2 or lifeNum > 3:
                        arr[row, col] = 0
                if a[row, col] == 0:
                    if lifeNum == 3:
                        arr[row, col] = 1

    print('第%d代：\n'%time, arr)
    print('还有{0}个细胞存活'.format(np.sum(arr)))
if __name__ == '__main__':
    lifeGame(10)
