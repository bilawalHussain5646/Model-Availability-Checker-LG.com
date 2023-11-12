import time
from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import tkinter as tk
import tkinter.font as tkFont
import threading
from selenium.webdriver.chrome.options import Options



def InfiniteScrolling(driver):
        last_height = driver.execute_script("return document.body.scrollHeight")
        while True:
            # Scroll down to bottom
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

            # Wait to load page
            time.sleep(4)

            # Calculate new scroll height and compare with last scroll height
            new_height = driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                break
            last_height = new_height



def LG_Web(driver,list_of_models,country):
        output_df = pd.DataFrame(columns=['Model','LG',"link"])
        
        for models in list_of_models:
                
                    print("Model: ",models)
                    issueAppear = True
                    while (issueAppear == True):
                        if country == "sa_en":
                            driver.get(f"https://www.lg.com/{country}/search/search-all?type=B2C&bizType=B2C&siteType=MKT&adobeSearchType=gnb&searchResultFlag=Y&search={models}&obsBuynowFlag=N")
                            time.sleep(5)
                            try:
                                sa_prod = driver.find_element(By.CSS_SELECTOR,"div[class='cs-search-result__all-flagbox']").text
                                sa_link = driver.find_element(By.CSS_SELECTOR,"a[class='title c-text-contents__headline']").get_attribute("href")
                                
                                print(sa_prod)
                                if models in sa_prod:
                                    issueAppear= False
                                    output_df = output_df.append({
                                                "Model":models,
                                                "LG": "o",
                                                "link": sa_link,
                                    },ignore_index=True)
                                    print(models,"Found")
                                else:
                                    issueAppear= False
                                    output_df = output_df.append({
                                            "Model":models,
                                            "LG": "x",
                                            "link": ""
                    
                                    },ignore_index=True)
                                    print(models,"Not Found")
                                
                            except:
                                print("error")
                                issueAppear = True
                        else:
                            driver.get(f"https://www.lg.com/{country}/search/search-all?type=B2C&bizType=B2C&siteType=MKT&adobeSearchType=gnb&searchResultFlag=Y&search={models}&obsBuynowFlag=N")
                    
                            time.sleep(5)
                            try:
                                ids = driver.find_element(By.CSS_SELECTOR,".list-box")
                                all_divs  = ids.find_elements(By.TAG_NAME, "li")
                                total_models = len(all_divs)
                    
                                counter = 0
                                for div in all_divs:
                                    try:
                                        title = div.find_element(By.CSS_SELECTOR,".sku")
                                        link = div.find_element(By.TAG_NAME,"a").get_attribute("href")
                                        print(link)
                                        title_value = title.text
                                        model_id = title_value
                                        print(model_id)
                                        if model_id.find(models) != -1:
                                            issueAppear= False
                                            output_df = output_df.append({
                                                "Model":models,
                                                "LG": "o",
                                                "link": link,
                                            },ignore_index=True)
                                            print(models,"Found")
                                            break
                                    except:
                                        pass
                                    issueAppear= False
                                    counter+=1

                                if counter == total_models:
                                    issueAppear= False
                                    output_df = output_df.append({
                                            "Model":models,
                                            "LG": "x",
                                            "link": ""
                    
                                    },ignore_index=True)
                                    print(models,"Not Found")
                            except:
                                issueAppear = True
                


                

        with pd.ExcelWriter("output.xlsx",mode="a",if_sheet_exists='replace') as writer:
            output_df.to_excel(writer,sheet_name=f"{country}")
        


def Run_LG():

    data = pd.read_excel("testing.xlsx",sheet_name="countries")
    
    driver = webdriver.Chrome("C:\Program Files (x86)\chromedriver.exe")
    list_of_countries = data["country"]

    for country in list_of_countries:
        country_sheet = pd.read_excel("testing.xlsx",sheet_name=f"{country}")
        list_of_models = country_sheet["Models"]
        LG_Web(driver,list_of_models,country)
    
  

                 
         

# Main App 
class App:

    def __init__(self, root):
        #setting title
        root.title("LG Models Check")
        ft = tkFont.Font(family='Arial Narrow',size=13)
        #setting window size
        width=640
        height=480
        screenwidth = root.winfo_screenwidth()
        screenheight = root.winfo_screenheight()
        alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        root.geometry(alignstr)
        root.resizable(width=False, height=False)
        root.configure(bg='black')

        ClickBtnLabel=tk.Label(root)
       
      
        
        ClickBtnLabel["font"] = ft
        
        ClickBtnLabel["justify"] = "center"
        ClickBtnLabel["text"] = "LG Model Check"
        ClickBtnLabel["bg"] = "black"
        ClickBtnLabel["fg"] = "white"
        ClickBtnLabel.place(x=120,y=190,width=150,height=70)
    

        
        Lulu=tk.Button(root)
        Lulu["anchor"] = "center"
        Lulu["bg"] = "#009841"
        Lulu["borderwidth"] = "0px"
        
        Lulu["font"] = ft
        Lulu["fg"] = "#ffffff"
        Lulu["justify"] = "center"
        Lulu["text"] = "START"
        Lulu["relief"] = "raised"
        Lulu.place(x=375,y=190,width=150,height=70)
        Lulu["command"] = self.start_func




  

    def ClickRun(self):

        running_actions = [
            Run_LG,          
         
        ]

        thread_list = [threading.Thread(target=func) for func in running_actions]

        # start all the threads
        for thread in thread_list:
            thread.start()

        # wait for all the threads to complete
        for thread in thread_list:
            thread.join()
    
    def start_func(self):
        thread = threading.Thread(target=self.ClickRun)
        thread.start()

    
        

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()


# Run()
