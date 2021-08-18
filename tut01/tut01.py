def meraki_helper(n):
    s = str(n)
    if len(s)==1:
        return True
    for i in range(len(s)-1):
        if abs(int(s[i])-int(s[i+1]))!=1:
            return False
    return True

if __name__ == "__main__":
    input_list = [12, 14, 56, 78, 98, 54, 678, 134, 789, 0, 7, 5, 123, 45, 76345,987654321]
    merakis = 0
    non_merakis = 0
    for ele in input_list:
        ans = meraki_helper(ele)
        if ans:
            print("Yes - "+str(ele)+" is a Meraki number")
            merakis += 1
        else:
            print("No - "+str(ele)+" is Not a meraki number")
            non_merakis += 1
    print("the input list contains "+str(merakis)+" meraki and "+str(non_merakis)+" non meraki numbers")
