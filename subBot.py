import tagui as r

r.init(visual_automation = True)
r.click('D:\SMWSED\edgelogo.png')
r.wait(3)
r.click('D:\SMWSED\googlePage.png')
r.wait(2)
r.type('/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input','amazon.in[enter]')
print("website is true ")
r.wait(2)
r.ask('Do you want to submit')
r.close()   