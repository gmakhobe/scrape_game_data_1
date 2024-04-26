import time
import datetime
import pandas as pd
import unittest
from selenium import webdriver
from selenium.webdriver.common.by import By
#from appium.webdriver.common.touch_action import TouchAction
from selenium.webdriver.common.action_chains import ActionChains
from openpyxl import Workbook



class InitialBookInformation():
  """
  
    Class to get urls from spreadsheet
  
  """

  def getBookName(self):
    currentDate = datetime.datetime.now()

    currentYear = currentDate.strftime("%Y")
    currentDay = currentDate.strftime("%d")
    currentMonth = currentDate.strftime("%m")

    return "Book_" + currentMonth + "_" + currentDay + "_" + currentYear + ".xlsx"

  def getBookData(self):
    dataframe = pd.read_excel("input/" + self.getBookName(), sheet_name="Sheet1")
    listOfTeamData = []
  
    rowsOfLinks = len(dataframe)
    
    count = 0
    while (count < rowsOfLinks):
      teamData = {
        "team_a": dataframe.iloc[count][0],
        "team_b": dataframe.iloc[count][1],
        "match_total": dataframe.iloc[count][2]
      }
      listOfTeamData.append(teamData)
      count = count + 1

    return listOfTeamData


class TeamInformation():
  """
  
    Class to generate team information
  
  """
  
  def __init__(self, team_url):
    self.driver = webdriver.Firefox()
    
    
    self.team_url = team_url
    self.team_name = None
    self.team_score_list = None

  def main(self):
    """ Main Function
      This function generates Team Name & Team Score List.
    """
    
    driver = self.driver
    
    driver.get(self.team_url)
    driver.implicitly_wait(3)
    
    isWebPageValid = self.isWebPageValid()
    if (isWebPageValid == False):
      print("Failed to load page!")
      return False
    
    time.sleep(3)
    self.acceptCookies()
    time.sleep(5)

    isTeamNameReceived = self.getTeamName()
    if isTeamNameReceived == False:
      print("Failed to get team name!")
      return False
    
    
    isTeamScoresReceived = self.getTeamScores()
    if isTeamScoresReceived == False:
      print("Failed to get team scores!")
      return False
    
    print(self.team_score_list)
    print("Done")
    time.sleep(4)
    return True

  def isWebPageValid(self):
    # 1. Validate that page has loaded
    try:
      self.driver.find_element(By.XPATH, "//div[@id='user-menu']/span[@class='header__text header__text--user']")
      return True
    except:
      return False

  def acceptCookies(self):
    # 2. Accept cookies
    try:
      accept_button = self.driver.find_element(By.XPATH, "//div[@id='onetrust-button-group']/button[@id='onetrust-accept-btn-handler']")

      time.sleep(2)
      accept_button.click()
      return True
    except:
      return False

  def getTeamName(self):
    # 3. Scroll down to find name of the team
    try:
      team_name = self.driver.find_element(By.XPATH, "//div[@class='heading__title']/div[@class='heading__name']")
    
      time.sleep(2)
    
      actions = ActionChains(self.driver)
    
      actions.move_to_element(team_name)
      actions.perform()
    except:
      return False
    time.sleep(2)
    
    self.team_name =  team_name.text
    return True
  
  def getTeamScores(self):
    # 4. Scroll down to find team scores
    team_score_list = []

    count = 0
    loop_counter = 0
    while (True):

      try:
        is_todays_match = True if "Today's Matches" == self.driver.find_element(By.XPATH, "//section[@class='ui-section team-page-summary event ui-section--topIndented']/h1[@class='ui-section__title']").text else False
      except:
        is_todays_match = False
      
      next_table = ""

      if is_todays_match:
        next_table = "section[@class='ui-section event ui-section--topIndented'][1]/"
        
      try:
        team_name = self.driver.find_element(By.XPATH, "//" + next_table + "div[@class='leagues--static event--leagues sportName basketball']/div[" + str(2 + loop_counter) + "]/div[3]").text
        team_score = None
      
        if team_name == self.team_name:
          team_score = self.driver.find_element(By.XPATH, "//" + next_table + "div[@class='leagues--static event--leagues sportName basketball']/div[" + str(2 + loop_counter) + "]/div[5]").text
        else:
          team_score = self.driver.find_element(By.XPATH, "//" + next_table + "div[@class='leagues--static event--leagues sportName basketball']/div[" + str(2 + loop_counter) + "]/div[6]").text
        
        loop_counter = loop_counter + 1
      except:
        loop_counter = loop_counter + 1
        continue

      print("Team score")
      print(team_score)
        
      team_score_list.append(int(team_score))

      if count == 5:
        break
      count = count + 1  
    
    self.team_score_list = team_score_list
    return True

  def __del__(self):
    self.driver.close()


class PostBookInformation():
  def __init__(self, match_data):
    self.match_data = match_data
    self.number_of_books = len(match_data)
  
  def createABook(self):
    """ Create new work book"""
    currentDate = datetime.datetime.now()

    currentYear = currentDate.strftime("%Y")
    currentDay = currentDate.strftime("%d")
    currentMonth = currentDate.strftime("%m")

    count = 0
    while (count < self.number_of_books):
      work_book = Workbook()
      work_book_sheet = work_book.active
      output = "Output-Book_" + currentMonth + "_" + currentDay + "_" + currentYear + "_" + str(count) + " " + self.match_data[count]["team_names"]["team_a"] +".xlsx"
      
      work_book_sheet["A2"] = "Date"
      work_book_sheet["B2"] = currentDay + "/" + currentMonth + "/" + currentYear
      work_book_sheet["A3"] = "Match"
      work_book_sheet["B3"] = self.match_data[count]["team_names"]["team_a"]
      work_book_sheet["C3"] = "vs"
      work_book_sheet["D3"] = self.match_data[count]["team_names"]["team_b"]
      work_book_sheet["A6"] = "=B3"
      work_book_sheet["B6"] = "=D3"
      work_book_sheet["C6"] = "Score Per Quarter"
      work_book_sheet["A7"] = 0
      work_book_sheet["B7"] = 0
      work_book_sheet["C7"] = "=SUM(A7;B7)"
      work_book_sheet["A8"] = 0
      work_book_sheet["B8"] = 0
      work_book_sheet["C8"] = "=SUM(A8;B8)"
      work_book_sheet["A9"] = 0
      work_book_sheet["B9"] = 0
      work_book_sheet["C9"] = "=SUM(A9;B9)"
      work_book_sheet["A10"] = 0
      work_book_sheet["B10"] = 0
      work_book_sheet["C10"] = "=SUM(A9;B9)"
      work_book_sheet["C11"] = "=SUM(C7;C8;C9;C10)"
      work_book_sheet["A12"] = "Expected Close (Under)"
      work_book_sheet["B12"] = self.match_data[count]["match_total"]
      work_book_sheet["D12"] = "Expected Close (Over)"
      work_book_sheet["E12"] = self.match_data[count]["match_total"]
      work_book_sheet["A13"] = "Q4 Average Close"
      work_book_sheet["B13"] = "=(SUM(C7:C8)/2)*4"
      work_book_sheet["D13"] = "Q4 Average Close"
      work_book_sheet["E13"] = "=(SUM(C7:C8)/2)*4"
      work_book_sheet["A14"] = "Q4 Close based on Q2"
      work_book_sheet["B14"] = "=SUM(C7:C8)*2"
      work_book_sheet["D14"] = "Q4 Close based on Q2"
      work_book_sheet["E14"] = "=SUM(C7:C8)*2"
      work_book_sheet["A15"] = "Q4 Close"
      work_book_sheet["B15"] = "=SUM(C7:C10)"
      work_book_sheet["D15"] = "Q4 Close"
      work_book_sheet["E15"] = "=SUM(C10)"
      work_book_sheet["A16"] = "Outcome"
      work_book_sheet["B16"] = '=IF(B12=0,"NA",IF(B15<B12,"Win", "Fail"))'
      work_book_sheet["D16"] = "Outcome"
      work_book_sheet["E16"] = '=IF(E12=0,"NA",IF(E15>E12,"Win", "Fail"))'
      work_book_sheet["B20"] = "=B3"
      work_book_sheet["C20"] = "=D3"
      work_book_sheet["A19"] = "Data"
      work_book_sheet["A21"] = "Highest Score"
      work_book_sheet["B21"] = max(self.match_data[count]["team_a_score"])
      work_book_sheet["C21"] = max(self.match_data[count]["team_b_score"])
      work_book_sheet["D21"] = "=SUM(B21:C21)"
      work_book_sheet["A22"] = "Lowest Score"
      work_book_sheet["B22"] = min(self.match_data[count]["team_a_score"])
      work_book_sheet["C22"] = min(self.match_data[count]["team_b_score"])
      work_book_sheet["D22"] = "=SUM(B22:C22)"
      work_book_sheet["A23"] =  "Average Score"
      work_book_sheet["B23"] =  sum(self.match_data[count]["team_a_score"])/len(self.match_data[count]["team_a_score"])
      work_book_sheet["C23"] =  sum(self.match_data[count]["team_b_score"])/len(self.match_data[count]["team_b_score"])
      work_book_sheet["D23"] =  "=SUM(B23,C23)"

      work_book_sheet["A25"] =  "=B3"
      work_book_sheet["A26"] =  self.match_data[count]["team_a_score"][0]
      work_book_sheet["A27"] =  self.match_data[count]["team_a_score"][1]
      work_book_sheet["A28"] =  self.match_data[count]["team_a_score"][2]
      work_book_sheet["A29"] =  self.match_data[count]["team_a_score"][3]
      work_book_sheet["A30"] =  self.match_data[count]["team_a_score"][4]
      work_book_sheet["A31"] =  self.match_data[count]["team_a_score"][5]

      work_book_sheet["B25"] =  "=D3"
      work_book_sheet["B26"] =  self.match_data[count]["team_b_score"][0]
      work_book_sheet["B27"] =  self.match_data[count]["team_b_score"][1]
      work_book_sheet["B28"] =  self.match_data[count]["team_b_score"][2]
      work_book_sheet["B29"] =  self.match_data[count]["team_b_score"][3]
      work_book_sheet["B30"] =  self.match_data[count]["team_b_score"][4]
      work_book_sheet["B31"] =  self.match_data[count]["team_b_score"][5]

      work_book.save(output)
      work_book_sheet = None
      work_book = None
      count = count + 1
    print("Done creating work books")


class App():
  def __init__(self):
    #init
    self.initBookInfo = InitialBookInformation()
    self.listOfTeamUrls = self.initBookInfo.getBookData()
    self.totalUrls = len(self.listOfTeamUrls)
    self.teamInformation = []
  
  def main(self):
    count = 0
    while (count < self.totalUrls):
      match_data = {
        "team_names": {
          "team_a": None,
          "team_b": None
        },
        "team_a_score": None,
        "team_b_score": None,
        "match_total": None
      }
      team_a_obj = TeamInformation(self.listOfTeamUrls[count]["team_a"])
      if team_a_obj.main() == False:
        print("Failed to get Team A information.")
        return -1
      
      match_data["team_names"]["team_a"] = team_a_obj.team_name
      match_data["team_a_score"] = team_a_obj.team_score_list
      match_data["match_total"] = self.listOfTeamUrls[count]["match_total"]
      
      team_a_obj = None # De-references
      
      team_b_obj = TeamInformation(self.listOfTeamUrls[count]["team_b"])
      if team_b_obj.main() == False:
        print("Failed to get Team B information.")
        return -1
      
      match_data["team_names"]["team_b"] = team_b_obj.team_name
      match_data["team_b_score"] = team_b_obj.team_score_list

      team_b_obj = None # De-references
      self.teamInformation.append(match_data)

      count = count + 1
    
    print(self.teamInformation)
    work_book = PostBookInformation(self.teamInformation)
    work_book.createABook()
    print("Done working")


if __name__ == "__main__":
  App().main()