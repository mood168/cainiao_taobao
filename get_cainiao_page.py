import json
import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import time
import traceback
import re
from datetime import datetime
import os

# 設置Chrome選項
chrome_options = Options()
chrome_options.add_argument('--ignore-certificate-errors')  # 忽略SSL證書錯誤
chrome_options.add_argument('--ignore-ssl-errors')  # 忽略SSL錯誤
chrome_options.add_argument('--window-size=1920,1080')  # 設置窗口大小
chrome_options.add_argument('--disable-infobars')  # 禁用infobars
chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])
chrome_options.add_experimental_option('useAutomationExtension', False)
# chrome_options.add_argument('--headless')  # 無頭模式,調試時先註釋掉

driver = None
try:
    print("開始執行自動化流程...")
    
    # 初始化WebDriver時添加選項
    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, 30)  # 增加默認等待時間
    print("瀏覽器已啟動")
    
    driver.get('https://desk.cainiao.com/')
    print("已訪問初始頁面")

    def retry_operation(operation, max_retries=5, delay=2):
        for i in range(max_retries):
            try:
                return operation()
            except Exception as e:
                if i == max_retries - 1:
                    raise e
                print(f"操作失敗,重試中... ({i+1}/{max_retries})")
                time.sleep(delay)

    # 等待並切換到iframe
    def switch_to_iframe():
        iframe = wait.until(
            EC.frame_to_be_available_and_switch_to_it((By.ID, 'alibaba-login-box'))
        )
        print("已切換到登入iframe")
        return iframe

    retry_operation(switch_to_iframe)

    # 輸入登入信息
    def input_credentials():
        username_input = wait.until(
            EC.presence_of_element_located((By.NAME, 'fm-login-id'))
        )
        username_input.clear()
        username_input.send_keys('Abby.Chen@presco.ws')
        print("已輸入用戶名")

        password_input = wait.until(
            EC.presence_of_element_located((By.NAME, 'fm-login-password'))
        )
        password_input.clear()
        password_input.send_keys('Abby1023')
        print("已輸入密碼")

        login_button = wait.until(
            EC.element_to_be_clickable((By.CLASS_NAME, 'fm-button'))
        )
        login_button.click()
        print("已點擊登入按鈕")

    retry_operation(input_credentials)

    # 切換回主文檔
    driver.switch_to.default_content()
    print("已切換回主文檔")
    
    # 等待登入完成
    print("等待登入完成...")
    
    def wait_for_login_success():
        def check_login():
            current_url = driver.current_url
            print(f"當前URL: {current_url}")
            if "login" not in current_url and "desk.cainiao.com" in current_url:
                print("登入成功!")
                return True
            return False

        max_retries = 2
        for i in range(max_retries):
            try:
                if check_login():
                    return True
                time.sleep(3)
                print(f"等待登入完成... ({i+1}/{max_retries})")
            except Exception as e:
                print(f"檢查登入狀態時發生錯誤: {e}")
            
        return False

    if not wait_for_login_success():
        raise Exception("登入超時")

    # 導向到任務頁面
    target_url = 'https://desk.cainiao.com/unified/myTask/pendingTask'
    print(f"正在導向到任務頁面: {target_url}")
    
    def navigate_to_task_page():
        driver.get(target_url)
        wait.until(lambda d: d.execute_script('return document.readyState') == 'complete')
        # 等待表格元素出現
        wait.until(EC.presence_of_element_located((By.TAG_NAME, 'table')))
        time.sleep(5)  # 額外等待動態內容

    retry_operation(navigate_to_task_page, max_retries=3, delay=5)
    
    print("檢查工作狀態...")
    
    def check_and_switch_status():
        # 等待狀態按鈕出現
        status_button = wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, 'span.Status--statusTrigger--1En5pHk')
        ))
        
        # 檢查狀態文字
        status_text = status_button.text.strip()
        print(f"當前狀態: {status_text}")
        
        if status_text == "下班":
            print("檢測到下班狀態,準備切換到上班...")
            
            # 移動滑鼠到狀態按鈕
            actions = webdriver.ActionChains(driver)
            actions.move_to_element(status_button).perform()
            print("已移動滑鼠到狀態按鈕")
            
            # 等待彈出選單出現
            wait.until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, 'div.coneProtal-overlay-wrapper.opened')
            ))
            print("狀態選單已出現")
            
            try:
                # 嘗試方法1: 通過 Status--statusItem--3UvMvXq 類別定位
                online_option = wait.until(EC.element_to_be_clickable((
                    By.XPATH, 
                    "//span[contains(@class, 'Status--statusItem--3UvMvXq') and contains(text(), '上班')]"
                )))
            except:
                try:
                    # 嘗試方法2: 通過父元素定位
                    online_option = wait.until(EC.element_to_be_clickable((
                        By.XPATH,
                        "//div[contains(@class, 'coneProtal-overlay-wrapper')]//span[contains(text(), '上班')]"
                    )))
                except:
                    # 嘗試方法3: 最寬鬆的定位方式
                    online_option = wait.until(EC.element_to_be_clickable((
                        By.XPATH,
                        "//*[contains(text(), '上班')]"
                    )))
            
            # 使用 JavaScript 點擊,避免可能的覆蓋問題
            driver.execute_script("arguments[0].click();", online_option)
            print("已點擊上班選項")
            
            # 等待狀態更新
            time.sleep(2)  # 等待狀態切換
            
            # 等待列表重新載入
            wait.until(EC.presence_of_element_located((By.TAG_NAME, 'table')))
            time.sleep(2)  # 額外等待確保列表完全載入
            
            # 驗證狀態是否已更新
            new_status = wait.until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, 'span.Status--statusTrigger--1En5pHk')
            )).text.strip()
            
            if new_status == "上班":
                print("已成功切換到上班狀態")
            else:
                raise Exception(f"狀態切換失敗,當前狀態: {new_status}")
        else:
            print("當前已是上班狀態,無需切換")
    
    # 重試切換狀態操作
    retry_operation(check_and_switch_status, max_retries=3, delay=2)
    
    # 確保列表已完全載入
    wait.until(EC.presence_of_element_located((By.TAG_NAME, 'table')))
    time.sleep(3)  # 額外等待確保列表穩定
    
    print("開始尋找工單號連結...")
    
    def find_order_links():
        # 等待表格完全加載
        wait.until(EC.presence_of_element_located((By.TAG_NAME, 'table')))
        
        # 使用多種定位策略
        order_links = None
        try:
            # 方法1: 使用XPath精確匹配14位數字
            order_links = driver.find_elements(By.XPATH, "//table//a[string-length(text())=14 and translate(text(), '0123456789', '') = '']")
        except:
            try:
                # 方法2: 使用CSS選擇器
                order_links = driver.find_elements(By.CSS_SELECTOR, "table tbody tr td:first-child a")
            except:
                # 方法3: 使用更寬鬆的XPath
                order_links = driver.find_elements(By.XPATH, "//*[contains(@class, 'ticket-number') or contains(@class, 'order-number')]")
        
        if not order_links:
            raise Exception("找不到工單號連結")
        print(f"找到 {len(order_links)} 個工單號連結")
        return order_links

    def load_processed_orders():
        try:
            if os.path.exists('processed_orders.json'):
                with open('processed_orders.json', 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    print("\n已處理的工單列表:")
                    for order_id, info in data.items():
                        print(f"工單號: {order_id}, 處理時間: {info['processed_time']}")
                    return data
            else:
                print("\n未找到處理記錄文件,將創建新的記錄")
                return {}
        except Exception as e:
            print(f"讀取處理記錄時發生錯誤: {str(e)}")
            print("將重新創建處理記錄")
            return {}

    def save_processed_order(order_id, order_url):
        try:
            processed_orders = load_processed_orders()
            if order_id not in processed_orders:  # 只有在不存在時才添加
                processed_orders[order_id] = {
                    'url': order_url,
                    'processed_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                }
                with open('processed_orders.json', 'w', encoding='utf-8') as f:
                    json.dump(processed_orders, f, ensure_ascii=False, indent=2)
                print(f"已將工單 {order_id} 添加到處理記錄")
            else:
                print(f"工單 {order_id} 已存在於處理記錄中")
        except Exception as e:
            print(f"保存處理記錄時發生錯誤: {str(e)}")
            print("錯誤詳情:", traceback.format_exc())

    # 獲取所有工單連結
    order_links = retry_operation(find_order_links, max_retries=10, delay=3)
    total_orders = len(order_links)
    print(f"總共有 {total_orders} 個工單需要處理")


    # 載入已處理的工單記錄
    processed_orders = load_processed_orders()
    print(f"已載入處理記錄,共 {len(processed_orders)} 個工單")

    # 檢查並建立records資料夾
    if not os.path.exists('records'):
        os.makedirs('records')
        print("已建立records資料夾")

    # 遍歷處理每個工單
    for index, order_link in enumerate(order_links, 1):
        current_order_id = None
        try:
            current_order_id = order_link.text
            order_url = order_link.get_attribute('href')
            
            # 檢查是否已處理過
            if current_order_id in processed_orders:
                print(f"\n工單 {current_order_id} 已於 {processed_orders[current_order_id]['processed_time']} 處理過")
                print(f"跳過處理,繼續下一個工單")
                continue
                
            print(f"\n開始處理第 {index}/{total_orders} 個工單: {current_order_id}")
            print(f"工單URL: {order_url}")
            
            # 在開始處理前先保存到記錄中,避免重複處理
            save_processed_order(current_order_id, order_url)
            
            # 點擊工單連結並切換到新窗口
            original_window = driver.current_window_handle
            old_handles = driver.window_handles
            
            # 使用JavaScript點擊
            driver.execute_script("arguments[0].click();", order_link)
            print("已點擊工單連結")
            
            try:
                # 等待新窗口出現
                wait.until(lambda d: len(d.window_handles) > len(old_handles))
                new_handle = [h for h in driver.window_handles if h not in old_handles][0]
                print(f"切換到新窗口: {new_handle}")
                
                # 切換到新窗口
                driver.switch_to.window(new_handle)
                
                try:
                    # 等待頁面加載完成
                    wait.until(lambda d: d.execute_script('return document.readyState') == 'complete')
                    time.sleep(5)  # 增加等待時間
                    
                    # 確保在"工单基本信息"標籤頁
                    info_tab = wait.until(EC.presence_of_element_located(
                        (By.XPATH, "//div[@class='next-tabs-tab-inner' and contains(text(), '工单基本信息')]")
                    ))
                    print("找到工单基本信息標籤")
                    
                    # 檢查是否需要點擊標籤
                    parent_tab = info_tab.find_element(By.XPATH, "./..")
                    if 'active' not in parent_tab.get_attribute('class'):
                        info_tab.click()
                        print("點擊工单基本信息標籤")
                        time.sleep(2)
                    
                    print("已切換到工单基本信息標籤頁")
                    
                    # 等待內容區域加載
                    wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'next-tabs-content')))
                    time.sleep(2)  # 額外等待內容加載
                    
                    # 抓取工單信息
                    field_data = {}
                    
                    # 等待並獲取所有字段行
                    rows = wait.until(EC.presence_of_all_elements_located(
                        (By.CSS_SELECTOR, '.next-row')
                    ))
                    print(f"找到 {len(rows)} 行數據")
                    
                    # 創建文件並寫入工單號
                    with open(f'records/{current_order_id}.txt', 'w', encoding='utf-8') as f:
                        f.write(f'工单号: {current_order_id}\n')
                        
                        # 遍歷每一行尋找指定字段
                        for row in rows:
                            try:
                                cols = row.find_elements(By.CSS_SELECTOR, '.next-col')
                                for i in range(0, len(cols), 2):
                                    if i + 1 < len(cols):
                                        field_name = cols[i].text.strip()
                                        field_value = cols[i + 1].text.strip()
                                        if '工单基本信息' in field_name:
                                            time_match = re.search(r'\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}', field_name)
                                            # 使用正則表達式匹配時間格式 YYYY-MM-DD hh:mm:ss
                                            if time_match:
                                                f.write(f'创建时间: {time_match.group()}\n')
                                            num_parts = field_name.split('运单号')
                                            if len(num_parts) > 1:
                                                num_parts_txt = num_parts[1].strip()
                                                f.write(f'运单号: {num_parts_txt[:8]}\n')
                                                shipmentNo = num_parts_txt[:8]
                                            description_parts = field_name.split('工单描述')
                                            if len(description_parts) > 1:
                                                f.write(f'\n\n工单描述: {description_parts[1]}\n')
                                        if '当前工单记录' in field_name:
                                            record_parts = field_name.split('当前工单记录')
                                            if len(record_parts) > 1:
                                                # 計算當前工單記錄中的時間戳數量
                                                all_matches = re.finditer(r'\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}', record_parts[1])
                                                match_count = len(list(all_matches))
                                                # 找到所有時間戳並添加換行符
                                                for post_time_match in re.finditer(r'\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}', record_parts[1]):
                                                    record_parts[1] = record_parts[1].replace(post_time_match.group(), post_time_match.group() + "\n")
                                                    
                                                f.write(f'\n\n当前工单记录: {record_parts[1]}\n')

                            except Exception as e:
                                print(f"處理行數據時發生錯誤: {str(e)}")
                                continue            
                    # 獲取 API token
                    try:
                        token_url = "https://ecapi.sp88.tw/api/Token"
                        token_headers = {
                            "Content-Type": "application/json",
                        }
                        token_data = {
                            "account": "ESCS",
                            "password": "SDG3jdkd59@1"
                        }
                        token_data_str = json.dumps(token_data)
                        token = None  # 初始化 token 變量
                        token_success = False  # 添加標誌變量
                        
                        print(f"正在請求 token, URL: {token_url}")
                        print(f"請求數據: {token_data_str}")
                        
                        try:
                            token_response = requests.post(token_url, headers=token_headers, data=token_data_str, timeout=30)
                            print(f"API 響應狀態碼: {token_response.status_code}")
                            print(f"API 響應內容: {token_response.text}")
                            
                            if token_response.status_code == 200:
                                token_data = token_response.json()
                                if token_data.get("token") != None:
                                    token = token_data.get("token")
                                    token_success = True  # 設置成功標誌
                                    print("成功獲取 token")
                                else:
                                    error_msg = token_data.get("Message", "未知錯誤")
                                    print(f"獲取 token 失敗,API 返回錯誤: {error_msg}")
                                    print(f"完整響應: {token_data}")
                            else:
                                print(f"HTTP 請求失敗,狀態碼: {token_response.status_code}")
                                print(f"錯誤響應: {token_response.text}")
                        except requests.exceptions.Timeout:
                            print("請求超時,請檢查網絡連接或 API 服務器狀態")
                        except requests.exceptions.RequestException as e:
                            print(f"請求異常: {str(e)}")
                        except json.JSONDecodeError as e:
                            print(f"解析 JSON 響應失敗: {str(e)}")
                            print(f"原始響應內容: {token_response.text}")
                        
                        # 只有在成功獲取 token 時才繼續執行
                        if token_success:
                            # 使用獲取的 token 調用追蹤 API
                            track_url = f"https://ecapi.sp88.tw/api/track/B2C?eshopId=74A&shipmentNo={shipmentNo}"  # 修正字符串插值語法
                            track_headers = {
                                "Content-Type": "application/json", 
                                "Authorization": f"Bearer {token}"
                            }
                            
                            print(f"正在請求追蹤信息, URL: {track_url}")
                            track_response = None
                            track_response = requests.get(track_url, headers=track_headers)
                            print(f"追蹤 API 響應狀態碼: {track_response.status_code}")
                            print(f"追蹤 API 響應內容: {track_response.text}")
                            
                            if track_response.status_code == 200:
                                track_data = track_response.json()
                                if track_data.get("status") != None:
                                    tracking_info = track_data.get("status")
                                    memo_date = datetime.strptime(track_data.get("date"), "%Y-%m-%d %H:%M:%S").strftime("%Y-%m-%d %H:%M:%S")
                                    errorCodeDescription = track_data.get("errorCodeDescription")
                                    errorCode = track_data.get("errorCode")
                                    # 將追蹤信息寫入文件
                                    with open(f'records/{current_order_id}.txt', 'a', encoding='utf-8') as f:
                                        f.write(f'(自動寫入)新留言: {tracking_info}\n{memo_date}')
                                        f.write(f'\n\n貨態查詢:')
                                        if errorCode == 0:
                                            f.write(f'\n結果: {track_response.text.replace('"', '').replace('\n', '')}\n')
                                        else:
                                            f.write(f'\n結果: {errorCodeDescription.replace('"', '')}\n')
                                    
                                    # 添加留言
                                    try:
                                        print("準備添加留言...")
                                        # 點擊留言按鈕
                                        leave_message_btn = wait.until(EC.element_to_be_clickable(
                                            (By.XPATH, "//span[contains(@class, 'next-btn-helper') and contains(text(), '留言')]")
                                        ))
                                        driver.execute_script("arguments[0].click();", leave_message_btn)
                                        print("已點擊留言按鈕")
                                        
                                        # 等待留言對話框出現
                                        wait.until(EC.presence_of_element_located(
                                            (By.CSS_SELECTOR, 'div.leaveMessage-overlay-wrapper.opened')
                                        ))
                                        print("留言對話框已出現")
                                        
                                        # 找到並填寫留言框
                                        message_textarea = wait.until(EC.presence_of_element_located(
                                            (By.CSS_SELECTOR, 'textarea[name="memo"]')
                                        ))
                                        message_textarea.clear()
                                        message_textarea.send_keys(str(tracking_info))
                                        print("已填寫留言內容")
                                        
                                        # 點擊確認按鈕
                                        confirm_btn = wait.until(EC.element_to_be_clickable(
                                            (By.XPATH, "//span[contains(@class, 'leaveMessage-btn-helper') and contains(text(), '确认')]")
                                        ))
                                        driver.execute_script("arguments[0].click();", confirm_btn)
                                        print("已點擊確認按鈕")
                                        
                                        # 等待留言對話框消失
                                        wait.until_not(EC.presence_of_element_located(
                                            (By.CSS_SELECTOR, 'div.leaveMessage-overlay-wrapper.opened')
                                        ))
                                        print("留言已成功送出")
                                        
                                    except Exception as e:
                                        print(f"添加留言時發生錯誤: {str(e)}")
                                        print("繼續處理其他步驟...")
                    except Exception as e:
                        print(f"調用 API 時發生錯誤: {str(e)}")
                    print(f"已完成數據寫入到 {current_order_id}.txt")
                    
                except Exception as e:
                    print(f"處理工單內容時發生錯誤: {str(e)}")
                    raise  # 重新拋出異常以觸發外層的清理代碼
                    
            except Exception as e:
                print(f"處理新窗口時發生錯誤: {str(e)}")
                raise  # 重新拋出異常以觸發外層的清理代碼
                
            # 在成功處理後保存記錄
            try:
                save_processed_order(current_order_id, order_url)
                print(f"已記錄工單 {current_order_id} 的處理狀態")
            except Exception as e:
                print(f"保存工單處理記錄時發生錯誤: {str(e)}")
                
        except Exception as e:
            print(f"處理工單 {current_order_id or '未知'} 時發生錯誤: {str(e)}")
            if driver.current_window_handle != original_window:
                try:
                    driver.close()
                    driver.switch_to.window(original_window)
                except Exception as close_error:
                    print(f"關閉錯誤窗口時發生異常: {str(close_error)}")
        else:
            # 正常完成時的清理代碼
            try:
                driver.close()
                driver.switch_to.window(original_window)
                print(f"已完成工單 {current_order_id} 的處理並關閉窗口")
            except Exception as close_error:
                print(f"關閉窗口時發生錯誤: {str(close_error)}")
        finally:
            time.sleep(2)  # 無論成功與否都等待一下再處理下一個
            
    print("\n所有工單處理完成!")

    def switch_to_offline():
        try:
            print("\n準備切換到下班狀態...")
            # 等待狀態按鈕出現
            status_button = wait.until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, 'span.Status--statusTrigger--1En5pHk')
            ))
            
            # 移動滑鼠到狀態按鈕
            actions = webdriver.ActionChains(driver)
            actions.move_to_element(status_button).perform()
            print("已移動滑鼠到狀態按鈕")
            
            # 等待彈出選單出現
            wait.until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, 'div.coneProtal-overlay-wrapper.opened')
            ))
            print("狀態選單已出現")
            
            try:
                # 嘗試方法1: 通過 Status--statusItem--3UvMvXq 類別定位
                offline_option = wait.until(EC.element_to_be_clickable((
                    By.XPATH, 
                    "//span[contains(@class, 'Status--statusItem--3UvMvXq') and contains(text(), '下班')]"
                )))
            except:
                try:
                    # 嘗試方法2: 通過父元素定位
                    offline_option = wait.until(EC.element_to_be_clickable((
                        By.XPATH,
                        "//div[contains(@class, 'coneProtal-overlay-wrapper')]//span[contains(text(), '下班')]"
                    )))
                except:
                    # 嘗試方法3: 最寬鬆的定位方式
                    offline_option = wait.until(EC.element_to_be_clickable((
                        By.XPATH,
                        "//*[contains(text(), '下班')]"
                    )))
            
            # 使用 JavaScript 點擊,避免可能的覆蓋問題
            driver.execute_script("arguments[0].click();", offline_option)
            print("已點擊下班選項")
            
            # 等待狀態更新
            time.sleep(2)  # 等待狀態切換
            
            # 驗證狀態是否已更新
            new_status = wait.until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, 'span.Status--statusTrigger--1En5pHk')
            )).text.strip()
            
            if new_status == "下班":
                print("已成功切換到下班狀態")
            else:
                raise Exception(f"狀態切換失敗,當前狀態: {new_status}")
                
        except Exception as e:
            print(f"切換到下班狀態時發生錯誤: {str(e)}")
            raise

    # 切換到下班狀態
    retry_operation(switch_to_offline, max_retries=3, delay=2)
    
    # 等待2秒
    time.sleep(2)
    print("準備關閉瀏覽器...")
    
    # 關閉瀏覽器
    driver.quit()
    print("瀏覽器已關閉")
    
except KeyboardInterrupt:
    print("\n檢測到Ctrl+C,正在優雅退出...")
except Exception as e:
    print(f"發生錯誤: {e}")
    print("錯誤詳情:")
    traceback.print_exc()
finally:
    if driver:
        driver.quit()
        print("瀏覽器已關閉")