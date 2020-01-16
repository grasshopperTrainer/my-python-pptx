def solution(n):
    import math
    def find_sub_num(v):
        nums = []
        # if v == 1:
        #     return [1]

        while True:
            is_prime = True
            for ii in range(2, v + 1):
                if v % ii == 0:
                    v = int(v / ii)
                    nums.append(ii)
                    is_prime = False
                    break
            if is_prime:
                nums.append(v)
                break
            if v == 1:
                break
        return nums

    answer = 0
    for yn in range(math.floor(n / 2) + 1):
        xn = n - yn * 2
        big = max((xn, yn))
        small = min((xn, yn))

        # print(f'big {big}, small {small}')
        if any([x == 0 for x in (big, small)]):
            answer += 1
            continue

        heads = []
        bottoms = []
        for v in range(yn + xn, big, -1):
            # print(f'checking {v}')
            a = find_sub_num(v)
            s = 1
            for i in a:
                s = s*i
            if s != v:
                raise
            print(f'{v} is mul of {a}')
            heads += a

        for v in range(2, small + 1):
            for i in find_sub_num(v):
                try:
                    heads.remove(i)
                except:
                    bottoms.append(i)
        head = 1
        bottom = 1
        for i in heads:
            head *= i
        for i in bottoms:
            bottom *= i

        try:
            answer += head / bottom
        except:
            s = 0
            for i in heads:
                s = s % 1234567
                s += i
            # print(s)
            s = s / bottom
            answer += s
            answer = answer % 1234567
    return int(answer % 1234567)

print(solution(2000))