import requests
import airbnb
import pandas as pd

# Tour Data
tour_place = None
adults_tourist = 0
tourist = list()
check_in = None
check_out = None

# Input tour Data
tourist = list()
while True:
    tourist = list()
    tour_place = input("Please input your tour place: ")
    adults_tourist = int(input("Please input your tourists number: "))
    check_in = input("Please input check in date (ex: 2020-01-01): ")
    check_out = input("Please input check out date (ex: 2020-01-01): ")
    for person in range(adults_tourist):
        tourist.append(input("Please input tourist {}'s name: ".format(person + 1)))
    print("Tour Place: {}\nTotal Tourist: {}\nTourist'names are {}\nCheck in date: {}\nCheck out date: {} ".format(tour_place, adults_tourist,tourist,check_in,check_out))
    if input("Confirm your request (Y/N): ") == "Y":
        break
# Excel Auto colume fit funciton
def get_col_widths(dataframe):
    idx_max = max([len(str(s)) for s in dataframe.index.values] + [len(str(dataframe.index.name))])
    return [idx_max] + [max([len(str(s)) for s in dataframe[col].values] + [len(col)]) for col in dataframe.columns]

# Airbnb Account Data
my_account = "thomson861106@gmail.com"
my_password = "b05502037"

# Initialize Airbnb API Function
API = airbnb.Api(username=my_account, password=my_password, randomize=True, access_token="cy3qna5ox65l3blxu96xko1ba")

# Request rooms data from get_home function
homes = API.get_homes(query=tour_place, checkin=check_in, checkout=check_out, items_per_grid=50, offset=0, adults=adults_tourist)
gethomes = homes['explore_tabs'][0]['sections']

# Get all avaliable rooms data
items = 0
listing_homes = gethomes[len(gethomes) - 1]['listings']
while homes['explore_tabs'][0]['pagination_metadata']['has_next_page'] is True:
    items += 50
    homes = API.get_homes(query=tour_place, checkin=check_in, checkout=check_out, items_per_grid=50, offset=items, adults=adults_tourist)
    gethomes = homes['explore_tabs'][0]['sections']
    if list(gethomes[len(gethomes) - 1].keys()).count('listings'):
        listing_homes += (gethomes[len(gethomes) - 1]['listings'])
    else:
        break

# Filter the data we wanted to dict
rooms = {'name': list(), 'total_price': list(), 'price/person': list(), 'person_capacity': list(), 'rate': list(), 'reviews_count': list()}
dict_url = dict()
n=0
for room in listing_homes:
    n+=1
    print(n)
    room_url = 'https://www.airbnb.com.tw/rooms/' + str(room['listing']['id']) + "?adults={}&check_in={}&check_out={}".format(adults_tourist, check_in, check_out)
    dict_url.update({room['listing']['name']:room_url})
    rooms['name'].append(room['listing']['name'])
    rooms['total_price'].append(room['pricing_quote']['price']['total']['amount'])
    rooms['price/person'].append(room['pricing_quote']['price']['total']['amount'] / adults_tourist)
    rooms['person_capacity'].append(room['listing']['person_capacity'])
    rooms['reviews_count'].append(room['listing']['reviews_count'])
    
    if list(room['listing'].keys()).count('avg_rating') == 1:
        rooms['rate'].append(room['listing']['avg_rating'])
    else:
        rooms['rate'].append(0)

# Bulid up a xlsx file to output datas
writer = pd.ExcelWriter(check_in+tour_place + ".xlsx", engine='xlsxwriter')

# Setting special format require
data_format = writer.book.add_format({'align': 'center'})
currency_format = writer.book.add_format({'align': 'center', 'num_format': '[$$-409]#,##0.00'})

# Writing all rooms data
worksheet = writer.book.add_worksheet('Rooms sheet')
worksheet.write(0, 0, 'index')
all_kind = list(rooms.keys())
for kind in range(len(all_kind)):
    worksheet.write(0, 1 + kind, all_kind[kind])   
for i in range(len(listing_homes)):
    worksheet.write(i + 1, 0, str(i + 1))
    worksheet.write_url(i + 1, 1, dict_url[rooms['name'][i]], string=rooms['name'][i])
    worksheet.write(i + 1, 2,rooms['total_price'][i])
    worksheet.write(i + 1, 3,rooms['price/person'][i])
    worksheet.write(i + 1, 4,rooms['person_capacity'][i])
    worksheet.write(i + 1, 5,rooms['rate'][i])
    worksheet.write(i + 1, 6,rooms['reviews_count'][i])


# Construct DataFrame to work on Auto Fit columns function above
rooms_data = pd.DataFrame(rooms)
for i, width in enumerate(get_col_widths(rooms_data)):
    worksheet.set_column(i, i, width+2,data_format)
worksheet.set_column("C:D", 15, currency_format)

# Input users voting require
total_rooms = len(listing_homes)
max_price = max(rooms['price/person'])
min_price = min(rooms['price/person'])
print("After request from airbnb API, {} rooms have been found.\nThe highest price per person is {} TWD, and the lowest price is {} TWD".format(total_rooms, max_price, min_price))
max_accept = int(input("What's your highest acceptable price per person: "))
show_number = int(input("How many rooms would yo like to show in Voting sheet: "))


# Filter the data required
out_rooms = rooms_data[rooms_data.loc[:,'price/person'] < max_accept]
out_rooms = out_rooms.sort_values(by=['rate','reviews_count'], ascending=False)
if show_number > out_rooms['name'].count():
    show_number = out_rooms['name'].count()
vote_rooms = out_rooms.iloc[range(show_number)]

# Write Voting sheet
vote_rooms.to_excel(writer, sheet_name='Vote sheet')
worksheet = writer.book.get_worksheet_by_name('Vote sheet')
for i, width in enumerate(get_col_widths(vote_rooms)):
    worksheet.set_column(i, i, width+2,data_format)
worksheet.write(0, 0, 'Rank',data_format)
worksheet.write(0, len(all_kind) + 2, 'Vote Sum')
worksheet.set_column(0, len(all_kind) + 2, 10, data_format)
worksheet.set_column("C:D", 15, currency_format)
for person in range(adults_tourist):
    worksheet.write(0 ,len(all_kind) + 3 + person, tourist[person], data_format)
for rank in range(show_number):
    worksheet.write(rank + 1, 0, rank + 1)
    worksheet.write_url(rank + 1, 1, dict_url[vote_rooms.iloc[rank][0]], string=vote_rooms.iloc[rank][0])
    worksheet.write(rank + 1, len(all_kind) + 2, "=SUM({}{}:{}{})".format(chr(65 + len(all_kind) + 3), rank + 2, 'ZZ', rank + 2), data_format)
all_kind = list(rooms.keys())
for kind in range(len(all_kind)):
    worksheet.write(0, 1 + kind, all_kind[kind])

# Write Budget Control System
worksheet = writer.book.add_worksheet("Budget Control")
headline=['Name', 'Content', 'Payment', 'Remarks']
for item in range(len(headline)):
    worksheet.write(0, item, headline[item])
    worksheet.set_column(0, item, len(headline[item]) + 2, data_format)
worksheet.set_column("C:C", 15, currency_format)
headline=['Name', 'Paid', 'Feedback']
for item in range(len(headline)):
    worksheet.write(0, item + 6, headline[item])
    worksheet.set_column(0, item + 6, len(headline[item]) + 2, data_format)
for person in range(adults_tourist):
    worksheet.write(1 + person, 6, tourist[person])
    worksheet.write(1 + person, 7, "=SUMIF(A:A,\"={}\",C:C)".format(tourist[person]))
    worksheet.write(1 + person, 8, "=H{}-SUM(C:C)/{}".format(person + 2, adults_tourist))
worksheet.set_column("H:I", 15, currency_format)
writer.save()

# End of this program
print("Finished. Have a fun trip^^")