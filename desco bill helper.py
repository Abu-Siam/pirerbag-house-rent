meter_index = {"3836": "6south", "3823": "6north", "9004": "5south", "2403": "5north", "5248": "4south",
               "9022": "4north",
               "9043": "3south", "4322": "3north", "4726": "2south", "3832": "2north", "3825": "1south",
               "9065": "1north",
               "9005": "pump", "3960": "checkmeter_total"}

bill_index = {"6south": 0, "6north": 0, "5south": 0, "5north": 0, "4south": 0, "4north": 0, "3south": 0, "3north": 0,
              "2south": 0, "2north": 0, "1south": 0, "1north": 0, "pump": 0, "checkmeter_total" : 0}
loop_end_flag=1
while(1):

    meter_input = input("give the meter index =")
    #bill_input = input("give the respective desco bill=")
    temp = meter_index[meter_input] + "=" + "bill"
    bill_index[meter_index[meter_input]]=0
   #print("output checking :", bill_index[meter_index[meter_input]])
    print("output checking :",bill_index[meter_index[meter_input]])
    print(temp)
    """
    while(1):
        confirmation=input("do you wanna try again?(y/n) ")
        if(confirmation=="y"):
            loop_end_flag=1
            break
        elif(confirmation=="n"):
            loop_end_flag=0
            break
        else:print("TYPE ERROR!!!!!!")"""