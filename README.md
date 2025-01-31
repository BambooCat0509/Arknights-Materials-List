本程式僅做為 "明日方舟 素材一覽.xlsx" 的配套計算使用

usage: "Arknights Materials Update.exe" [-h ] [-d ] [-f ] [-n [ ...]] [-m ] [-c ]

options:
  --Help, -H, -h        show this help message and exit
  --ChromeDriverPath , --Driver , -D , -d
                        Chrome Driver 的絕對路徑
  --FilePath , --File , -F , -f
                        "明日方舟 素材一覽.xlsx" 的絕對路徑
  --NoCount [ ...], --NC [ ...], -N [ ...], -n [ ...]
                        不想列入計算的關卡代號 / 活動名稱 (簡體字或英文，多個以空格區分，ex: -n 长夜临光 CW)
  --Minimun , --Min , -M , -m
                        是否獲取素材單件最低期望理智 (True: 1, False: 0)
  --Comprehensive , --Com , -C , -c
                        是否計算綜合素材最高效率關卡 (True: 1, False: 0)