ERROR_COUNTER = 0
ERROR_SMALL = 0

for i in range(10):
    try:
        try:
            if i % 2 == 0:
                raise Exception
        except Exception as inst:
            ERROR_SMALL += 1
            print(f"De kleine error counter staat op: {ERROR_SMALL}")    
    except Exception as inst:
        ERROR_COUNTER += 1
        print(f"De grote error counter staat op: {ERROR_COUNTER}")
