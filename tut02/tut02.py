import queue
def get_memory_score(arr):
    # in queue we can remove elements from front and add elements to back
    q= queue.Queue()
    x = set()
    points = 0
    for i in arr:
        if i in x:
            points+=1
        else:
            if len(x) ==5:
                val = q.get()
                x.remove(val)
            q.put(i)
            x.add(i)
    return points

input_nums = [3, 4, 1, 6, 3, 3, 9,0, 0,0]
# converting values to strings as isdigit works only for strings
input_nums = [str(i) for i in input_nums]
# to store invalid elements if any
invalid = []
# to store valid elements is any
valid = []
for i in input_nums:
    if i.isdigit():
        valid.append(i)
    else:
        invalid.append(i)
# if we found any invalid elements then we print them
if len(invalid)>0:
    print('Please enter a valid input list. Invalid inputs detected:', invalid)
else:
    print( ' Score :' , get_memory_score(valid))