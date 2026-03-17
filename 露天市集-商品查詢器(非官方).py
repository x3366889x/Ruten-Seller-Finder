import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import webbrowser
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import time
import sys
import os

#捲動清單設定
class ScrollableFrame(tk.Frame):
    #初始化設定
    def __init__(self, container, bg_color="#ccf2ff"):
        #建立一個有tk.Frame功能的視窗
        super().__init__(container)
        #使用內建樣式
        style = ttk.Style()
        style.theme_use('alt')
        style.configure("Vertical.TScrollbar",bordercolor=bg_color) #消除邊界線
        #建立一個清單放在視窗
        self.canvas = tk.Canvas(self, bg=bg_color, highlightthickness=0, bd=0)
        #在清單建立一個滾動條
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview, style="Vertical.TScrollbar")
        #在清單建立一個放輸入資料的地方
        self.scrollable_frame = tk.Frame(self.canvas, bg=bg_color)
        #建立一筆資料 使用導入的資訊
        self.scroll_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        #綁定滾動調和清單的位置關係
        self.canvas.configure(yscrollcommand=scrollbar.set)
        #把清單靠左 填滿整個區塊 大小可變動
        self.canvas.pack(side="left", fill="both", expand=True)
        #把滾動條放右側 垂直填滿整個清單
        scrollbar.pack(side="right", fill="y")
        #設定比例與最小寬度
        self.scrollable_frame.columnconfigure(0, weight=5, minsize=100)
        self.scrollable_frame.columnconfigure(1, weight=5, minsize=100)
        self.scrollable_frame.columnconfigure(2, weight=2, minsize=40)
        self.scrollable_frame.columnconfigure(3, weight=2, minsize=40)
        self.scrollable_frame.columnconfigure(4, weight=3, minsize=40)
        #當輸入資料的大小或內容變化時，觸發 _on_frame_configure 方法 更新滾動條大小
        self.scrollable_frame.bind("<Configure>", self._on_frame_configure)
        #當清單自身大小變化時，觸發 _on_canvas_configure 方法 更新輸入資料大小
        self.canvas.bind("<Configure>", self._on_canvas_configure)
    
    #滾動條大小修正
    def _on_frame_configure(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    #輸入資料大小修正
    def _on_canvas_configure(self, event):
        self.canvas.itemconfig(self.scroll_window, width=event.width)

#主體視窗設定
class ProductListApp:
    #視窗初始設定函數
    def __init__(self, root): 
        #視窗基本設定
        self.root = root
        self.root.title("露天市集 商品查詢器")
        #更換程式圖示
        try:
            import sys, os
            #自動判斷是直接執行還是打包後的執行
            if hasattr(sys, '_MEIPASS'):
                icon_path = os.path.join(sys._MEIPASS, "icon.ico")
            else:
                icon_path = os.path.abspath("icon.ico")
            self.root.iconbitmap(icon_path)
        except Exception as e:
            print(f"圖示載入失敗: {e}")
            pass
        self.root.geometry("800x400")
        root.resizable(True, True)
        #邊框和功能區塊
        tk.Frame(root, height=1, bg="#000000").pack(side="top", fill="x")
        border_frame = tk.Frame(root, background="#888888")
        border_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        main_frame = tk.Frame(border_frame)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=3, pady=3)
        #設定區塊縮放權重 分割功能區塊
        main_frame.columnconfigure(0, minsize=350)
        main_frame.columnconfigure(1, weight=0)
        main_frame.columnconfigure(2, weight=1)
        main_frame.rowconfigure(0, weight=1)
        #左區塊
        left_frame = tk.Frame(main_frame, bg="#ccf2ff")
        left_frame.grid(row=0, column=0, sticky="nsew")
        #右區塊
        right_frame = tk.Frame(main_frame, bg="#ffe0e6")
        right_frame.grid(row=0, column=2, sticky="nsew")
        #左右區塊分隔線
        separator = tk.Frame(main_frame, width=3, bg="#888888")
        separator.grid(row=0, column=1, sticky="ns")
        #設定左區塊分區
        left_frame.columnconfigure(0, weight=36)
        left_frame.columnconfigure(1, weight=36)
        left_frame.columnconfigure(2, weight=6)
        left_frame.columnconfigure(3, weight=6)
        left_frame.columnconfigure(4, weight=17)
        left_frame.rowconfigure(0, weight=0)
        left_frame.rowconfigure(1, weight=0)
        left_frame.rowconfigure(2, weight=1)
        left_frame.rowconfigure(3, weight=0)
        left_frame.rowconfigure(4, weight=0)
        #設定對應區塊標頭
        tk.Label(left_frame, bd=1, relief="solid",   text="包含字段").grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        tk.Label(left_frame, bd=1, relief="solid",   text="不含字段").grid(row=0, column=1, sticky="nsew", padx=5, pady=5)
        tk.Label(left_frame, bd=1, relief="solid",   text="庫存"    ).grid(row=0, column=2, sticky="nsew", padx=5, pady=5)
        tk.Label(left_frame, bd=1, relief="solid",   text="預算"    ).grid(row=0, column=3, sticky="nsew", padx=5, pady=5)
        tk.Label(left_frame, background="#ccf2ff", text="　　"    ).grid(row=0, column=4, sticky="nsew", padx=5, pady=5)
        #設定分隔線
        tk.Frame(left_frame, height=3, bg="#888888").grid(row=1, column=0, columnspan=5, sticky="ew")
        #設定捲動視窗
        self.scrollable = ScrollableFrame(left_frame, bg_color="#ccf2ff")
        self.scrollable.grid(row=2, column=0, columnspan=5, sticky="nsew")
        #設定分隔線
        tk.Frame(left_frame, height=3, bg="#888888").grid(row=3, column=0, columnspan=5, sticky="ew")
        #設定新增商品按鈕
        btn = tk.Button(left_frame, text="增加一筆商品", command=self.add_product_row)
        btn.grid(row=4, column=0, columnspan=5, sticky="nsew", padx=5, pady=5)
        #儲存數據的字典和編號
        self.rows = {}
        self.next_row = 0
        #設定右區塊分區
        right_frame.columnconfigure(0, weight=1)
        right_frame.rowconfigure(0, weight=0)
        right_frame.rowconfigure(1, weight=0)
        right_frame.rowconfigure(2, weight=0)
        right_frame.rowconfigure(3, weight=0)
        right_frame.rowconfigure(4, weight=1)
        #第一列輸入參數
        container1 = tk.Frame(right_frame, bg="#ffe0e6")
        container1.grid(row=0, column=0, sticky="ew")
        tk.Label(container1, text="搜尋加載等待上限", bg="#ffe0e6").pack(side="left", padx=5, pady=5)
        self.entry1_1 = tk.Entry(container1, width=5)
        self.entry1_1.pack(side="left", padx=5, pady=5)
        self.entry1_1.insert(0, "15")
        tk.Label(container1, text="秒 | 詳情確認等待上限", bg="#ffe0e6").pack(side="left", padx=5, pady=5)
        self.entry1_2 = tk.Entry(container1, width=5)
        self.entry1_2.pack(side="left", padx=5, pady=5)
        self.entry1_2.insert(0, "5")
        tk.Label(container1, text="秒", bg="#ffe0e6").pack(side="left", padx=5, pady=5)
        #第二列輸入參數
        container2 = tk.Frame(right_frame, bg="#ffe0e6")
        container2.grid(row=1, column=0, sticky="ew")
        tk.Label(container2, text="是否可使用運費卷", bg="#ffe0e6").pack(side="left", padx=5, pady=5)
        self.var1 = tk.BooleanVar()
        chk1 = tk.Checkbutton(container2, variable=self.var1, bg="#ffe0e6", activebackground="#ffe0e6", text=" | 爬蟲時是否將模擬視窗移置螢幕外")
        chk1.pack(side="left")
        self.var2 = tk.BooleanVar()
        self.var2.set(True)
        chk2 = tk.Checkbutton(container2, variable=self.var2, bg="#ffe0e6", activebackground="#ffe0e6")
        chk2.pack(side="left")
        #第三列功能按鍵
        container3 = tk.Frame(right_frame, bg="#ffe0e6")
        container3.grid(row=2, column=0, sticky="ew")
        self.button3_1 = tk.Button(container3, text="查詢", command=self.search_products)
        self.button3_1.pack(side="left", padx=5, pady=5)
        self.button3_2 = tk.Button(container3, text="清除結果", command=self.clear_results)
        self.button3_2.pack(side="left", padx=5, pady=5)
        tk.Label(container3, text="(請慎選條件 符合商品太多 會查到天荒地老)", bg="#ffe0e6").pack(side="left", padx=5, pady=5)
        #第四列分隔線
        tk.Frame(right_frame, height=3, bg="#888888").grid(row=3, column=0, sticky="ew")
        #第五列捲動清單 + 滾動條
        columns = ("store","total","url")
        tree_frame = tk.Frame(right_frame)
        tree_frame.grid(row=4, column=0, sticky="nsew", padx=5, pady=5)
        #建立 Treeview
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        self.tree.heading("store", text="店鋪名稱")
        self.tree.heading("total", text="總金額")
        self.tree.heading("url", text="店鋪網址")
        self.tree.column("store", anchor="center", width=50)
        self.tree.column("total", anchor="center", width=50)
        self.tree.column("url", anchor="center", width=50)
        self.tree.pack(side="left", fill="both", expand=True)
        #雙擊事件監聽
        self.tree.bind("<Double-1>", self.open_url_from_tree)
        #加上垂直滾動條
        #追加粉紅色樣式
        style = ttk.Style()
        style.configure("Pink.Vertical.TScrollbar", bordercolor="#ffe0e6")
        #使用設定使用粉色樣式(為更改部分將繼承預設物件Vertical.TScrollbar的設定)
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview, style="Pink.Vertical.TScrollbar")
        self.tree.configure(yscrollcommand=scrollbar.set)
        #把捲動條放右邊
        scrollbar.pack(side="right", fill="y")

    #點兩下網址跳轉
    def open_url_from_tree(self, event):
        item_id = self.tree.identify_row(event.y)
        if item_id:
            url = self.tree.item(item_id, "values")[2]
            webbrowser.open(url)

    #刪除商品項
    def delete_row(self, row):
        #刪除GUI
        for widget in self.rows[row]:
            widget.destroy()
        #刪除字典對應資料
        del self.rows[row]

    #新增商品項
    def add_product_row(self):
        #減化物件內全域變量名稱
        sf = self.scrollable.scrollable_frame
        r = self.next_row
        #設定5個物件
        e1 = tk.Entry(sf, width=1)
        e2 = tk.Entry(sf, width=1)
        e3 = tk.Entry(sf, width=1)
        e3.insert(0, "1")
        e4 = tk.Entry(sf, width=1)
        e4.insert(0, "9999")
        b5 = tk.Button(sf, text="刪除", command=lambda row=r: self.delete_row(row))
        #將5個物件顯示到捲動清單
        e1.grid(row=r, column=0, sticky="nsew", padx=2, pady=2)
        e2.grid(row=r, column=1, sticky="nsew", padx=2, pady=2)
        e3.grid(row=r, column=2, sticky="nsew", padx=2, pady=2)
        e4.grid(row=r, column=3, sticky="nsew", padx=2, pady=2)
        b5.grid(row=r, column=4, sticky="nsew", padx=2, pady=2)
        #將變數儲存到字典變量中
        self.rows[r] = (e1, e2, e3, e4, b5)
        self.next_row += 1

    #清除結果
    def clear_results(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
    
    #查詢
    def search_products(self):
        #商品數量檢測
        if len(self.rows) == 0:
            messagebox.showerror("提醒", "空氣是免費的 會呼吸就行")
            return
        # 檢查等待時間是否為有效的正整數
        try:
            wait_page = int(self.entry1_1.get())
            wait_seller = int(self.entry1_2.get())
            if wait_page <= 0 or wait_seller <= 0:
                raise ValueError
        except:
            messagebox.showerror("錯誤", "請輸入有效的正整數作為等待時間（秒）")
            return
        #取得所有商品條件
        query_list = []
        for row_id in self.rows:
            e1, e2, e3, e4, _ = self.rows[row_id]
            include = e1.get().strip()
            exclude = e2.get().strip()
            stock   = e3.get().strip()
            budget  = e4.get().strip()
            #判斷輸入是否正確
            if not include or not stock or not budget:
                messagebox.showerror("格式錯誤", f"商品『{include}』『包含字段』『庫存』『預算』不可為空！")
                return
            try:
                stock_val = int(stock)
                budget_val = int(budget)
                if stock_val <= 0 or budget_val <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror("格式錯誤", f"商品『{include}』庫存與預算必須是大於0的整數")
                return
            item = {
                "include": include,
                "exclude": exclude,
                "stock": stock,
                "budget": budget,
            }
            query_list.append(item)
        #設定時間基準
        start_time = time.time()

        #初始化視窗===================================================================================================

        #新增視窗
        options = Options()
        #模擬真實身分
        user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
        options.add_argument(f'user-agent={user_agent}')
        #抓完HTML架構就視為加載完
        options.page_load_strategy = 'eager'
        #不加載圖片
        prefs = {
            "profile.managed_default_content_settings.images": 2,
            "profile.default_content_settings.images": 2
        }
        options.add_experimental_option("prefs", prefs)
        #減少不必要渲染 減少資源占用
        options.add_argument("--disable-gpu")
        #防止JS執行速度變慢
        options.add_argument("--disable-background-timer-throttling")
        options.add_argument("--disable-backgrounding-occluded-windows")
        options.add_argument("--disable-renderer-backgrounding")
        options.add_argument("--memory-pressure-off")
        #移除 Chrome 的自動化標籤
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)
        #只顯示嚴重報錯
        options.add_argument("--log-level=3")
        #自動處理驅動程式版本
        service = Service(ChromeDriverManager().install())
        #啟動瀏覽器
        driver = webdriver.Chrome(service=service, options=options)
        #強制指定視窗寬高為全螢幕
        driver.set_window_size(
            self.root.winfo_screenwidth(),
            self.root.winfo_screenheight()
        )
        #設定視窗座標到螢幕外
        if self.var2.get():
            driver.set_window_position(self.root.winfo_screenwidth()-1, 0)
        else:
            driver.set_window_position(0, 0)

        #開始爬蟲=====================================================================================================

        #清除舊資料/顯示狀態/更新視窗
        self.clear_results()
        self.tree.insert("", "end", values=("初步尋找可能賣家","", ""))
        self.root.update()
        #生成賣家與A商品基本陣列
        self.list_list = []
        include = query_list[0]["include"]
        exclude = query_list[0]["exclude"]
        stock   = query_list[0]["stock"]
        budget  = query_list[0]["budget"]
        seller  = ""
        self.search(driver, include, exclude, stock, budget, seller)
        if len(query_list) > 1:
            #補其賣家其餘商品資訊
            for i, item in enumerate(self.list_list):
                self.clear_results()
                self.tree.insert("", "end", values=(f"正在判斷店家({i+1}/{len(self.list_list)})","", ""))
                self.root.update()
                for item2 in query_list[1:]:
                    include = item2["include"]
                    exclude = item2["exclude"]
                    stock   = item2["stock"]
                    budget  = item2["budget"]
                    seller  = item[0]
                    result = self.search(driver, include, exclude, stock, budget, seller)
                    self.list_list[i].append(result)
                    if result == False:
                        break
        #剔除有尋找失敗的商家
        self.list_list = [shop for shop in self.list_list if False not in shop]
        #釋放瀏覽器資源
        driver.quit()

        #輸出結果=====================================================================================================

        #如果沒有找到符合所有商品的共同賣家
        if self.list_list == []:
            self.clear_results()
            self.tree.insert("", "end", values=("無符合店家","", ""))
            return
        #整理出 店家 網址 總價
        result_list = []
        for item in self.list_list:
            x = 0
            for i, item2 in enumerate(query_list):
                x += item[i+2] * int(item2["stock"])
            result_list.append({
                "seller"    : item[1],
                "total"     : x,
                "seller_url": item[0]
            })
        #低到高排序
        result_list.sort(key=lambda x: x["total"])
        #投到清單
        self.clear_results()
        for result in result_list:
            self.tree.insert("", "end", values=(result["seller"], result["total"], result["seller_url"]))
        self.root.update()
        end_time = time.time()
        elapsed_time = end_time - start_time
        #結束提示
        minutes = int(elapsed_time // 60)
        seconds = int(elapsed_time % 60)
        time_str = f"{minutes} 分 {seconds} 秒" if minutes > 0 else f"{seconds} 秒"
        messagebox.showinfo("搜尋完成", f"總計用時：{time_str}")
        ###############################################################################################################
        print("-" * 30)
        print(f"目前名單總數: {len(self.list_list)}")
        for i, shop in enumerate(self.list_list):
            # 假設 shop[0] 是網址, shop[1] 是名稱, 其餘是價格
            print(f"編號 {i} | 賣家: {shop[0]} | 資料長度: {len(shop)}")
        print("-" * 30)
    
    #找商品
    def search(self, driver, include, exclude, stock, budget, seller):
        x = 0
        list_filter = []
        list_url    = []
        include_ = [w.strip() for w in include.split(" ") if w.strip()]
        exclude_ = [w.strip() for w in exclude.split(" ") if w.strip()]
        while(True):
            #翻頁
            x+=1
            #搜尋
            url0 = f'{seller if seller else "https://www.ruten.com.tw/"}find/?q={include}{"&isshippingdiscount=1" if self.var1.get() else ""}&prc.now=0-{budget}&p={x}&sort=prc%2Fac'
            driver.get(url0)
            #等待商品加載
            wait = WebDriverWait(driver, int(self.entry1_1.get()))
            try:
                wait.until(EC.any_of(
                    EC.visibility_of_element_located((By.CLASS_NAME, "rt-product-card-shopping-cart-action")),
                    lambda d: len(d.find_element(By.CSS_SELECTOR, "p.rt-mt-3x").text.strip()) > 0,
                    lambda d: len(d.find_element(By.CSS_SELECTOR, "div.store-search-alert").text.strip()) > 0
                ))
            except:
                messagebox.showerror("提示", f"搜尋『{include}』時：等待超時")
                return False
            #暫存網頁
            page_content = driver.page_source
            #全站搜尋沒結果
            if seller == "" and "查詢不到符合的商品" in page_content:
                if x==1:
                    messagebox.showerror("提示", f"搜尋結果無符合『{include}』的商品")
                    return
                else:
                    break
            #賣家店內沒結果
            if seller != "":
                try:
                    alert_box = driver.find_element(By.CSS_SELECTOR, "div.store-search-alert")
                    if "目前賣家沒有該商品" in alert_box.text:
                        if x == 1:
                            return False
                        else:
                            break
                except:
                    pass
            #滑動加載所有商品
            last_height = driver.execute_script("return document.body.scrollHeight")
            while True:
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                new_height = driver.execute_script("return document.body.scrollHeight")
                if new_height == last_height:
                    break
                last_height = new_height
            #把符合商品丟進陣列
            #儲存此頁所有商品
            product_blocks = driver.find_elements("css selector", ".rt-product-card-detail-wrap")
            for block in product_blocks:
                try:
                    link_element = block.find_element("css selector", "a.rt-product-card-name-wrap")
                    name = link_element.text.strip()
                    url1 = link_element.get_attribute("href")
                    #去除廣告或異常物件
                    if not name:
                        continue
                    #排除字段檢查
                    if (
                        any(keyword in name for keyword in exclude_)
                        or not all(keyword in name for keyword in include_)
                    ):
                        continue
                    #預算檢查
                    try:
                        price_element = block.find_element("css selector", ".rt-text-price")
                        price_raw = price_element.get_attribute("textContent").strip()
                        price = int(''.join(filter(str.isdigit, price_raw)))
                        if price > int(budget):
                            break
                    except Exception as e:
                        continue
                    list_filter.append([url1,price])
                except:
                    continue
        #檢查 同商家/庫存不足 & 補充資料
        for item in list_filter:
            driver.get(item[0])
            #等待商品加載(等到庫存有文字，且賣家名稱與網址都不是空的)
            wait = WebDriverWait(driver, int(self.entry1_2.get()))
            try:
                # 只檢查庫存文字與店名文字是否存在
                wait.until(lambda d: (
                    len(d.find_elements(By.CSS_SELECTOR, 'strong.rt-text-isolated')) > 0 and
                    len(d.find_element(By.CSS_SELECTOR, 'strong.rt-text-isolated').text.strip()) > 0 and
                    len(d.find_elements(By.CSS_SELECTOR, 'a.no-underline[title="查看賣場"]')) > 0 and
                    len(d.find_element(By.CSS_SELECTOR, 'a.no-underline[title="查看賣場"]').text.strip()) > 0
                ))
            except Exception as e:
                # 這裡建議印出具體錯誤，方便知道是哪邊超時
                print(f"⚠️ 偵測失敗，跳過連結商品{item[0]}。原因: {e}")
                continue
            #抓庫存數量
            try:
                #嘗試抓庫存數量
                stock_elem = driver.find_element("css selector", 'strong.rt-text-isolated')
                stock_text = stock_elem.text.strip()
                try:
                    actual_stock = int(''.join(filter(str.isdigit, stock_text)))
                except:
                    continue
            except:
                continue
            if actual_stock < int(stock):
                continue
            #抓賣家名稱與網址
            h2_elem = driver.find_element("css selector", 'h2.relative.flex.items-center.gap-3.text-lg.font-semibold')
            seller_elem = h2_elem.find_element("css selector", 'a[title="查看賣場"]')
            name = seller_elem.text.strip()
            url2 = seller_elem.get_attribute("href")
            #賣家查重
            if url2 in list_url:
                continue
            if seller != "":
                return item[1]
            else:
                list_url.append(url2)
                self.list_list.append([url2,name,item[1]])
                print(str([url2,name,item[1]]) + " | " + str(time.ctime()))
        return False

#啟動程式
root = tk.Tk()
app = ProductListApp(root)
root.mainloop()