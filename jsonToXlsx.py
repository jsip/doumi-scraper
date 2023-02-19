import os
from openpyxl import Workbook
import json

# loop through the listes directory and get every jSON file
for file in os.listdir('listes'):
    if file.endswith(".json"):
        # open the json file
        with open('listes/'+file) as json_file:
            # load the json file
            data = json.load(json_file)
            # create a workbook
            wb = Workbook()
            # get the active worksheet
            ws = wb.active
            # loop through the key "data"
            for row in data['data']:
              print(row)
              # select the keys [companyName, website, companyPhone, jobTitle, profileImageURL, phone, email, location["city"], email, freeMails[0] (if exists), name, personalEmail, ['socialUrls']['socialMedia'][0]['socialNetworkUrl'], positionStartDate]
              companyName = row['companyName'] if 'companyName' in row else ''
              website = row['website'] if 'website' in row else ''
              companyPhone = row['companyPhone'] if 'companyPhone' in row else ''
              jobTitle = row['jobTitle'] if 'jobTitle' in row else ''
              profileImageURL = row['profileImageURL'] if 'profileImageURL' in row else ''
              phone = row['phone'] if 'phone' in row else ''
              email = row['email'] if 'email' in row else ''
              city = row.get('location').get('city') if 'location' in row else ''
              freeMails = row['freeMails'][0] if 'freeMails' in row else ''
              name = row['name'] if 'name' in row else ''
              personalEmail = row['personalEmail'] if 'personalEmail' in row else ''
              positionStartDate = row['positionStartDate'] if 'positionStartDate' in row else ''

              # write the data to the worksheet
              ws.append([companyName, website, companyPhone, jobTitle, profileImageURL, phone, email, city, freeMails, name, personalEmail, positionStartDate])

            # save the workbook with the same name as the json file in the excel directory
            wb.save('excel/'+file.replace('.json', '.xlsx'))
