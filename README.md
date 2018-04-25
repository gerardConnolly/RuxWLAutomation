import pandas as pd
import numpy as np

import datetime
from datetime import date, timedelta

import sys
#sys.path.append('/opt/mnt/publicdrive/Analytics/Gerard/Utils/')
sys.path.append('/Volumes/ugcompanystorage/Company/public/Analytics/Gerard/Utils/GA/')
from SQL.Query import Query
from GA.GA_obj import GA

import openpyxl
from openpyxl.styles import colors
from openpyxl.styles import Font, Color, Border

def find_row(sheet):
    row = 1
    col = 1
    active_cell = 0

    while(active_cell != None):
        row = row + 1
        print(row)
        active_cell = (sheet.cell(row=row, column=col)).value
        print (active_cell)

    print('---')
    print(row)
    return row

def ga_query(start_date, end_date, filter_var, metrics, dimensions, max_results, segment, sort):

    ## The GA object takes  a profile ID and the location of your credential file as argument to create the object
    #External = GA('120725274', filelocation='/opt/mnt/publicdrive/Analytics/Gerard/Utils/GA/')
    External = GA('120725274', filelocation='/Volumes/ugcompanystorage/Company/public/Analytics/Gerard/Utils/GA/')
    query_response = External.get_results(start_date=start_date,
                    end_date=end_date,
                    filter_var=(None if filter_var==0 else filter_var),
                    metrics=(None if metrics==0 else metrics),
                    dimensions=(None if dimensions==0 else dimensions),
                    max_results=(None if max_results==0 else max_results),
                    segment=(None if segment==0 else segment),
                    sort=(None if sort==0 else sort))

    # Results are stored in 'rows'
    query_response['rows']

    # coverts list to a dataframe and grabs the first value from the dataframe
    result = pd.DataFrame(query_response['rows']).iloc[0, 0]

    # result:
    return result
#----------------------------------------------------------------------------------------------------------------------
today_date = pd.to_datetime('today')
today_date

# Today's date
today_date.weekday()

# Set __LAST_SUNDAY__ to today's date
__LAST_SUNDAY__ = today_date

# Until __LAST_SUNDAY__ is a Sunday (=6)
while (__LAST_SUNDAY__.weekday() != 6):
    __LAST_SUNDAY__ = __LAST_SUNDAY__ - datetime.timedelta(1)

print(__LAST_SUNDAY__)

# Calculate next monday date


__LAST_MONDAY__ = __LAST_SUNDAY__ - datetime.timedelta(6)
__LAST_MONDAY__

#output location

#---------------------------------------------------------------------------------------------------------------
#formatting Variables

percent_format = "##.##%"




#---------------------------------------------------------------------------------------------------------------
#set excel file variable to correct path and define the sheet variable
RuxWishList = openpyxl.load_workbook('/Users/gconnolly/Documents/projects/RUX/test.xlsx')
#RuxWishList = openpyxl.load_workbook('/Volumes/ugcompanystorage/Company/public/Analytics/Nikita/RUX_automated.xlsx')


Wish_list = RuxWishList.get_sheet_by_name('Wish_List')
empty_row = find_row(Wish_list)

#---------------------------------------------------------------------------------------------------------------
#Calculate th date/week and add it to column 1
col = 1
year = datetime.date.today().year
end_date = datetime.date.today()
week = end_date.isocalendar()[1] - 1
Wish_list.cell(row = empty_row, column = col).value = week
(Wish_list.cell(row = empty_row, column = col)).value = int("{}{}".format(year,week))
Wish_list.cell(row = empty_row, column = col).number_format = '#'
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)

#---------------------------------------------------------------------------------------------------------------
#Add the date range for the querys to follow into the excel sheet
col = col + 1

(Wish_list.cell(row = empty_row, column = col)).value = ("{} to {}, {}".format(__LAST_MONDAY__.strftime('%b %d'), __LAST_SUNDAY__.strftime('%b %d'), __LAST_SUNDAY__.strftime('%Y')))
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)


#---------------------------------------------------------------------------------------------------------------

# 1. Number of users that visit /basket

col = col+1

users_accessed_query = """
    SELECT count (distinct users.email)
                    from wish_list, users
                    where wish_list.accessed is not null
                    and wish_list.user_id = users.user_id
                    and TO_CHAR(wish_list.accessed,'YYYYMMDD') >= '{0}'
                    and TO_CHAR(wish_list.accessed,'YYYYMMDD') <= '{1}'
    """.format(__LAST_MONDAY__, __LAST_SUNDAY__)

users_accessed_df = Query('ugpostgres', users_accessed_query)
users_accessed = users_accessed_df.iloc[0][0]
users_accessed
print('\n1. Number of users that visit basket: '+ str(int(users_accessed)))

(Wish_list.cell(row=empty_row, column=col)).value = int(users_accessed)
Wish_list.cell(row=empty_row, column=col).number_format = '#,##0'
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)



#---------------------------------------------------------------------------------------------------------------
#2. GA Query for lists_viewed column D

col = col + 1

start_date = __LAST_MONDAY__.strftime('%Y-%m-%d')
end_date = __LAST_SUNDAY__.strftime('%Y-%m-%d')
filter_var = 'ga:pagePath=@/wish-list/my-wishlists?listId='
metrics = 'ga:pageviews'
dimensions = 0
max_results = 0
segment=0
sort=0

lists_viewed = ga_query(start_date, end_date, filter_var, metrics, dimensions, max_results, segment, sort)
print('\n1. GA QUERY Lists Viewed: '+ str(lists_viewed))

(Wish_list.cell (row=empty_row, column=col)).value = int(lists_viewed)
Wish_list.cell(row = empty_row, column=col).number_format = '#,##0'
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)




#---------------------------------------------------------------------------------------------------------------
#3. Users Added column E

col = col + 1

users_added_query = """
SELECT count(distinct users.email)
                 from wish_list_sku, wish_list, users
                 where wish_list_sku.wish_list_id = wish_list.id
                 and wish_list.user_id = users.user_id
                 and users.email !~* '(.*uncommongoods.*)'
                 and users.email !~* '(.*ugoods.*)'
                 and users.email !~* '(.*okeweka.*)'
                 and users.email !~* '(.*somethingsilly.*)'
                 and TO_CHAR(wish_list_sku.added,'YYYYMMDD') >= '{0}'
                 and TO_CHAR(wish_list_sku.added,'YYYYMMDD') <= '{1}'
""".format(__LAST_MONDAY__, __LAST_SUNDAY__)

users_added_df = Query('ugpostgres', users_added_query)
users_added = users_added_df.iloc[0][0]
users_added
print('\n2. Number of users that added a wishlist: ' + str(int(users_added)))
(Wish_list.cell(row=empty_row, column = col)).value = int(users_added)
Wish_list.cell(row=empty_row, column=col).number_format = '#,##0 '
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)





#---------------------------------------------------------------------------------------------------------------
#4. Users Added Loves column F
col = col + 1

users_added_to_loves_query = """
SELECT to_char(min(wish_list_sku.added),'yyyy-mm-dd') AS week_beginning, count(distinct users.email)
                from wish_list_sku, wish_list, users
                where wish_list_sku.wish_list_id = wish_list.id
                and wish_list.user_id = users.user_id
                and wish_list.name  ~* '(.*items i @heart@)'
                 and users.email !~* '(.*uncommongoods.*)'
                 and users.email !~* '(.*ugoods.*)'
                 and users.email !~* '(.*okeweka.*)'
                 and users.email !~* '(.*somethingsilly.*)'
                and TO_CHAR(wish_list_sku.added,'YYYYMMDD') >= '{0}'
                and TO_CHAR(wish_list_sku.added,'YYYYMMDD') <= '{1}'
""".format(__LAST_MONDAY__, __LAST_SUNDAY__)

users_added_to_loves_df = Query('ugpostgres', users_added_to_loves_query)
users_added_to_loves = users_added_to_loves_df.iloc[0][1]
users_added_to_loves

print('\n3. Number of users that added loves: ' + str(int(users_added_to_loves)))

(Wish_list.cell(row = empty_row, column = col)).value = int(users_added_to_loves)
Wish_list.cell(row=empty_row, column=col).number_format = '#,##0 '
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)

#---------------------------------------------------------------------------------------------------------------
#column space for calculation column G
col = col + 1

previous_row= empty_row - 1
one_year_ago = empty_row -53
formula_cell = (Wish_list.cell(row=previous_row, column=col)).value



#(Wish_list.cell(row = empty_row, column = col)).value = str(formula_cell)
(Wish_list.cell(row = empty_row, column = col)).value = "=F{}/F{}-1".format(empty_row,one_year_ago)
Wish_list.cell(row = empty_row, column = col).number_format = '##%'
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)
#---------------------------------------------------------------------------------------------------------------
#column H
col = col + 1

users_added_to_lists_query = """
SELECT to_char(min(wish_list_sku.added),'yyyy-mm-dd')AS week_beginning, count(distinct users.email)
                from wish_list_sku, wish_list, users
                where wish_list_sku.wish_list_id = wish_list.id
                and wish_list.user_id = users.user_id
                and wish_list.name !~* '(.*items i @heart@)'
                 and users.email !~* '(.*uncommongoods.*)'
                 and users.email !~* '(.*ugoods.*)'
                 and users.email !~* '(.*okeweka.*)'
                 and users.email !~* '(.*somethingsilly.*)'
                and TO_CHAR(wish_list_sku.added,'YYYYMMDD') >= '{0}'
                and TO_CHAR(wish_list_sku.added,'YYYYMMDD') <= '{1}'
""".format(__LAST_MONDAY__, __LAST_SUNDAY__)

users_added_to_lists_df = Query('ugpostgres', users_added_to_lists_query)
users_added_to_lists = users_added_to_lists_df.iloc[0] [1]
users_added_to_lists

print('\n4. Number of users that added lists: ' + str(int(users_added_to_lists)))

(Wish_list.cell(row = empty_row, column = col)).value = int(users_added_to_lists)
Wish_list.cell(row = empty_row, column = col).number_format = '#,##0 '
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)

#---------------------------------------------------------------------------------------------------------------
#column space for calculation column I
col = col + 1

previous_row= empty_row - 1
one_year_ago = empty_row -53
(Wish_list.cell(row = empty_row, column = col)).value = "=H{}/H{}-1".format(empty_row, one_year_ago)
Wish_list.cell(row = empty_row, column = col).number_format = percent_format
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)
#---------------------------------------------------------------------------------------------------------------
#column space for calculation column J (hybrid users)
col = col + 1

(Wish_list.cell(row = empty_row, column = col)).value = "=IFERROR(((F{}+H{})-E{})/E{},\"\")".format(empty_row,empty_row,empty_row,empty_row)
Wish_list.cell(row = empty_row, column = col).number_format = '##%'
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)
#---------------------------------------------------------------------------------------------------------------
#Column K

col = col + 1

new_users_added_query ="""SELECT
                     date_trunc('week', users.date_created) AS week_beginning,
                     count(distinct users.email) AS newusers
                FROM wish_list_sku, wish_list, users
                where 1=1
                and COALESCE(users.date_created::text, '') != ''
                and users.user_id = wish_list.user_id
                and wish_list.id = wish_list_sku.wish_list_id
                and wish_list_sku.added >= users.date_created
                and wish_list_sku.added <= date_trunc('week', users.date_created) + interval '7 day'  --????
                and date_trunc('week', users.date_created) is not null
                and TO_CHAR(users.date_created,'YYYYMMDD') >= '{0}'
                and TO_CHAR(users.date_created,'YYYYMMDD') <= '{1}'
            GROUP BY date_trunc('week', users.date_created)
            ORDER BY date_trunc('week', users.date_created) desc
""".format(__LAST_MONDAY__, __LAST_SUNDAY__)

new_users_added_df = Query('ugpostgres', new_users_added_query)
new_users_adding = new_users_added_df.iloc[0][1]
new_users_adding

print('\n5. Number of new users added: ' + str(int(new_users_adding)))

(Wish_list.cell(row = empty_row, column = col)).value = int(new_users_adding)
Wish_list.cell(row = empty_row, column = col).number_format = '#,##0 '
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)
#---------------------------------------------------------------------------------------------------------------
#column space for calculation column (returning users added Column L )
col = col + 1
(Wish_list.cell(row = empty_row, column = col)).value = "=IFERROR(E{}-K{},\"\")".format(empty_row,empty_row)
Wish_list.cell(row = empty_row, column = col).number_format = '#,##0 '
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)
#---------------------------------------------------------------------------------------------------------------
#column space for calculation column (returning user growth Column M)
col = col + 1
(Wish_list.cell(row = empty_row, column = col)).value = "=L{}/L{}-1".format(empty_row, one_year_ago)
Wish_list.cell(row = empty_row, column = col).number_format = '##%'
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)
#---------------------------------------------------------------------------------------------------------------
#column space for calculation column (new user growth column N)
col = col + 1

(Wish_list.cell(row = empty_row, column = col)).value = "=K{}/K{}-1".format(empty_row, one_year_ago)
Wish_list.cell(row = empty_row, column = col).number_format = '##%'
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)
#---------------------------------------------------------------------------------------------------------------
#column space for calculation column (returning % Column O)
col = col + 1
(Wish_list.cell(row = empty_row, column = col)).value = "=IFERROR(L{}/E{},\"\")".format(empty_row,empty_row)
Wish_list.cell(row = empty_row, column =col).number_format = '##%'
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)
#---------------------------------------------------------------------------------------------------------------
#column P
col = col + 1
new_users_added_to_loves_query = """SELECT date_trunc('week', users.date_created) AS week_beginning,
                count(distinct users.email) AS newusers
                FROM wish_list_sku, wish_list, users
                where date_trunc('week', users.date_created) is not null
                and users.user_id = wish_list.user_id
                and wish_list.id = wish_list_sku.wish_list_id
                and wish_list.name ~* '(.*items i @heart@.*)'
                and wish_list_sku.added >= users.date_created
                and wish_list_sku.added <= date_trunc('week', users.date_created) + interval '7 day'
                and date_trunc('week', users.date_created) is not null
                and TO_CHAR(users.date_created,'YYYYMMDD') >= '{0}'
                and TO_CHAR(users.date_created,'YYYYMMDD') <= '{1}'
                GROUP BY date_trunc('week', users.date_created)
                ORDER BY date_trunc('week', users.date_created)   desc
""".format(__LAST_MONDAY__, __LAST_SUNDAY__)

new_users_added_to_loves_df = Query('ugpostgres', new_users_added_to_loves_query)
new_user_loves = new_users_added_to_loves_df.iloc[0][1]
new_user_loves

print('\n6. Number of New Users Added To Loves: ' + str(int(new_user_loves)))

(Wish_list.cell(row = empty_row, column = col)).value = int(new_user_loves)
Wish_list.cell(row = empty_row, column = col).number_format = '#,##0 '
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)
#---------------------------------------------------------------------------------------------------------------
#column space for calculation column Q (unnamed)
col = col + 1
(Wish_list.cell(row = empty_row, column = col)).value = "=P{}/P{}-1".format(empty_row, one_year_ago)
Wish_list.cell(row = empty_row, column = col).number_format = '##%'
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)
#---------------------------------------------------------------------------------------------------------------
#column space for calculation column R (ret user loves)
col = col + 1
(Wish_list.cell(row = empty_row, column = col)).value = "=IFERROR(F{}-P{},\"\")".format(empty_row,empty_row)
Wish_list.cell(row = empty_row, column =col).number_format = '######### '
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)
#---------------------------------------------------------------------------------------------------------------
#column space for calculation column S(loves new user growth)
col = col + 1
(Wish_list.cell(row = empty_row, column = col)).value = "=P{}/P{}-1".format(empty_row, one_year_ago)
Wish_list.cell(row = empty_row, column = col).number_format = '##%'
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)
#---------------------------------------------------------------------------------------------------------------
#column space for calculation column T(loves ret user growth)
col = col + 1
(Wish_list.cell(row = empty_row, column = col)).value = "=R{}/R{}-1".format(empty_row, one_year_ago)
Wish_list.cell(row = empty_row, column = col).number_format = '##%'
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)
#---------------------------------------------------------------------------------------------------------------
#New User Lists Column U
col = col + 1
new_users_added_to_lists_query = """SELECT date_trunc('week', users.date_created) AS week_beginning,
                count(distinct users.email) AS newusers
                FROM wish_list_sku, wish_list, users
                where date_trunc('week', users.date_created) is not null
                and users.user_id = wish_list.user_id
                and wish_list.id = wish_list_sku.wish_list_id
                and wish_list.name !~* '(.*items i @heart@.*)'
                and wish_list_sku.added >= users.date_created
                and TO_CHAR(users.date_created,'YYYYMMDD') >= '{0}'
                and TO_CHAR(users.date_created,'YYYYMMDD') <= '{1}'
                and wish_list_sku.added <= date_trunc('week', users.date_created) + interval '7 day'
                GROUP BY date_trunc('week', users.date_created)
                ORDER BY date_trunc('week', users.date_created) desc

""".format(__LAST_MONDAY__, __LAST_SUNDAY__)

new_users_added_to_lists_df = Query('ugpostgres', new_users_added_to_lists_query)
new_users_lists = new_users_added_to_lists_df.iloc[0][1]
new_users_lists

print('\n7. Number of New Users list added: ' + str(int(new_users_lists)))

(Wish_list.cell(row = empty_row, column = col)).value = int(new_users_lists)
Wish_list.cell(row = empty_row, column = col).number_format = '#,##0 '
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)
#---------------------------------------------------------------------------------------------------------------
#column space for calculation column V ( new users list growth )
col = col + 1
(Wish_list.cell(row = empty_row, column = col)).value = "=U{}/U{}-1".format(empty_row, one_year_ago)
Wish_list.cell(row = empty_row, column = col).number_format = percent_format
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)
#---------------------------------------------------------------------------------------------------------------
#column space for calculation column W ( % of lovers that are returning )
col = col + 1

(Wish_list.cell(row = empty_row, column = col)).value = "=IFERROR((F{}-P{})/F{},\"\")".format(empty_row, empty_row, empty_row)
Wish_list.cell(row = empty_row, column = col).number_format = percent_format
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)
#---------------------------------------------------------------------------------------------------------------
#column space for calculation column X ( % of listers that are returning )
col = col + 1

(Wish_list.cell(row = empty_row, column =  col)).value = "=IFERROR((H{}-U{})/H{},"")".format(empty_row, empty_row, empty_row)
Wish_list.cell(row = empty_row, column = col).number_format = percent_format
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)
#---------------------------------------------------------------------------------------------------------------
#Column Y
col = col + 1

users_adding_dupe_skus_query = """SELECT
                       count (distinct USER_ID)
                    from
                       (
                          select
                             USER_SKU,
                             USER_ID,
                             count(USER_SKU) DUPES
                          from
                             (
                                select
                                   wish_list.user_id USER_ID,
                                   wish_list_sku.sku::text || wish_list.user_id::text USER_SKU
                                from
                                   wish_list_sku,
                                   wish_list
                                where
                                   wish_list.id = wish_list_sku.wish_list_id
                                   and TO_CHAR(wish_list_sku.added, 'YYYYMMDD') >= '{0}'
                                   and TO_CHAR(wish_list_sku.added, 'YYYYMMDD') <= '{1}'
                                group by
                                   wish_list_sku.sku::text || wish_list.user_id::text,
                                   wish_list.id,
                                   wish_list_sku.sku,
                                   wish_list.user_id
                             ) output1
                          group by
                             USER_SKU,
                             USER_ID
                       ) output2
                    where
                       DUPES >= 2 """.format(__LAST_MONDAY__, __LAST_SUNDAY__)

users_adding_dupe_skus_df = Query('ugpostgres', users_adding_dupe_skus_query)
users_adding_dupe_skus = users_adding_dupe_skus_df.iloc[0][0]
users_adding_dupe_skus

print('\n8. Users Adding Dupe SKUs: ' + str(int(users_adding_dupe_skus)))

(Wish_list.cell(row = empty_row, column = col)).value = int(users_adding_dupe_skus)
Wish_list.cell(row = empty_row, column = col).number_format = '#,##0 '
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)
#---------------------------------------------------------------------------------------------------------------
#column space for calculation column Z ( #percent returning, there is no calculation in here )
col = col + 1




#---------------------------------------------------------------------------------------------------------------
#column space for calculation column AA ( % of users adding dupes)
col = col + 1
(Wish_list.cell(row = empty_row, column = col)).value = "=IFERROR((F{}-P{})/F{},"")".format(empty_row, empty_row, empty_row)
Wish_list.cell(row = empty_row, column = col).number_format = percent_format
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)
#---------------------------------------------------------------------------------------------------------------
#Google Analytics Sessions query Column AB
col = col + 1

start_date = __LAST_MONDAY__.strftime('%Y-%m-%d')
end_date = __LAST_SUNDAY__.strftime('%Y-%m-%d')
filter_var = 0
metrics = 'ga:sessions'
dimensions = 0
max_results = 0
segment = 0

sort = 0

Total_sessions = ga_query(start_date, end_date, filter_var, metrics, dimensions, max_results, segment, sort)
print('\nGA2. Total sessions for the period: ' + Total_sessions)

(Wish_list.cell(row=empty_row, column=col)).value = int(Total_sessions)
Wish_list.cell(row=empty_row, column=col).number_format = '#,##0 '
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)
#---------------------------------------------------------------------------------------------------------------
#column space for calculation column AC( #user added / sessions calculations )
col = col + 1

(Wish_list.cell(row = empty_row, column = col)).value = "=IFERROR(E{}/AB{},\"\")".format(empty_row, empty_row)
Wish_list.cell(row = empty_row, column = col).number_format = '#.#####%'
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)
#---------------------------------------------------------------------------------------------------------------#
#column space for calculation column AD ( #user growth calculation )
col = col + 1

(Wish_list.cell(row = empty_row, column = col)).value = "=IFERROR(E{}/E{}-100%,\"\")".format(empty_row, one_year_ago)
Wish_list.cell(row = empty_row, column = col).number_format = percent_format
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)
# ---------------------------------------------------------------------------------------------------------------
#column AE
col = col + 1
wl_created_query = """
SELECT count(distinct wish_list.id)
                from wish_list, users
                where TO_CHAR(wish_list.created,'YYYYMMDD') >= '{0}'
                and TO_CHAR(wish_list.created,'YYYYMMDD') <= '{1}'
                and wish_list.user_id = users.user_id
                 and users.email !~* '(.*uncommongoods.*)'
                 and users.email !~* '(.*ugoods.*)'
                 and users.email !~* '(.*okeweka.*)'
                 and users.email !~* '(.*somethingsilly.*)'


""".format(__LAST_MONDAY__, __LAST_SUNDAY__)

wl_created_df = Query('ugpostgres', wl_created_query)
wl_created = wl_created_df.iloc[0][0]
wl_created

print('\n9. Wish Lists Created: ' + str(int(wl_created)))

(Wish_list.cell(row = empty_row, column = col)).value = int(wl_created)
Wish_list.cell(row = empty_row, column = col).number_format = '#,##0 '
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)

#---------------------------------------------------------------------------------------------------------------
#column space for calculation column AF ( #wishlist growth calculation )
col = col + 1
(Wish_list.cell(row = empty_row, column = col)).value = "=IFERROR(AE{}/AE{}-1,\"\")".format(empty_row, one_year_ago)
Wish_list.cell(row = empty_row, column = col).number_format = percent_format
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)
#---------------------------------------------------------------------------------------------------------------
#Column AG
col = col + 1

wl_name_created_query = """SELECT count(output.WLNAME) NAMED_LISTS
                    from (
                            select wish_list.name WLNAME, users.user_id USERID
                            from wish_list, users
                            where TO_CHAR(wish_list.created,'YYYYMMDD') >= '{0}'
                            and TO_CHAR(wish_list.created,'YYYYMMDD') <= '{1}'
                            and wish_list.user_id = users.user_id
                             and users.email !~* '(.*uncommongoods.*)'
                             and users.email !~* '(.*ugoods.*)'
                             and users.email !~* '(.*okeweka.*)'
                             and users.email !~* '(.*somethingsilly.*)'
                             and wish_list.name !~* '(.*items i @heart@)'
                            group by wish_list.name, users.user_id
                        ) output
""".format(__LAST_MONDAY__, __LAST_SUNDAY__)

wl_name_created_df = Query('ugpostgres', wl_name_created_query)
wl_name_created = wl_name_created_df.iloc [0] [0]
wl_name_created

print('\n10. Wish List Names Created: ' + str(int(wl_name_created)))

(Wish_list.cell(row = empty_row, column = col)).value = int(wl_name_created)
Wish_list.cell(row = empty_row, column = col).number_format = '#,##0'
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)
#---------------------------------------------------------------------------------------------------------------
#Column AH
col = col + 1
wl_empty_query = """
SELECT
                       count(output.WLID) EMPTY_LISTS
                    from
                       (
                          select
                             wish_list.id WLID,
                             users.user_id
                          from
                             users,
                             wish_list
                             left join
                                wish_list_sku
                                on wish_list_sku.wish_list_id = wish_list.id
                          where
                             TO_CHAR(wish_list.created, 'YYYYMMDD') >= '{0}'
                             and TO_CHAR(wish_list.created, 'YYYYMMDD') <= '{1}'
                             and wish_list.name !~* '(.*items i @heart@.*)'
                             and wish_list_sku.sku is null
                             and wish_list.user_id = users.user_id
                          group by
                             wish_list.id,
                             users.user_id
                       ) output
""".format(__LAST_MONDAY__, __LAST_SUNDAY__)
wl_empty_df = Query('ugpostgres', wl_empty_query)
wl_empty = wl_empty_df.iloc[0][0]

print('\n11. Empty Wish Lists: ' + str(int(wl_empty)))

(Wish_list.cell(row = empty_row, column = col)).value = int(wl_empty)
Wish_list.cell(row = empty_row, column = col).number_format = '#,##0'
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)
#---------------------------------------------------------------------------------------------------------------
#column AI
col = col + 1
skus_added_query = """
SELECT
                       count(output.SKU)
                    from
                       (
                          select
                             wish_list_sku.sku SKU,
                             wish_list.id WLID,
                             users.user_id
                          from
                             wish_list_sku,
                             wish_list,
                             users
                          where
                             TO_CHAR(wish_list_sku.added, 'YYYYMMDD') >= '{0}'
                             and TO_CHAR(wish_list_sku.added, 'YYYYMMDD') <= '{1}'
                             and wish_list_sku.wish_list_id = wish_list.id
                             and wish_list.user_id = users.user_id
                             and users.email !~* '(.*uncommongoods.*)'
                             and users.email !~* '(.*ugoods.*)'
                             and users.email !~* '(.*okeweka.*)'
                             and users.email !~* '(.*somethingsilly.*)'
                          group by
                             wish_list_sku.sku,
                             wish_list.id,
                             users.user_id
                       ) output

""".format(__LAST_MONDAY__, __LAST_SUNDAY__)
skus_added_df = Query('ugpostgres', skus_added_query)
skus_added = skus_added_df.iloc[0][0]

print('\n13. Skus Added to Wishlists: ' + str(int(skus_added)))

(Wish_list.cell(row = empty_row, column = col)).value = int(skus_added)
Wish_list.cell(row = empty_row, column = col).number_format = '#,##0 '
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)
#---------------------------------------------------------------------------------------------------------------
#column space for calculation column AJ( skus / list 7 days  hidden column)
col = col + 1

#---------------------------------------------------------------------------------------------------------------
#column space for calculation column AK ( skus / list 1 month  hidden column)
col = col + 1
#---------------------------------------------------------------------------------------------------------------
#column space for calculation column AL( percent empty)
col = col + 1

(Wish_list.cell(row = empty_row, column = col)).value = "=IFERROR(AH{}/AE{},\"\")".format(empty_row, empty_row)
Wish_list.cell(row = empty_row, column = col).number_format = percent_format
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)
#---------------------------------------------------------------------------------------------------------------
#column space for calculation column AM( sku add growth)
col = col + 1
(Wish_list.cell(row = empty_row, column = col)).value = "=IFERROR(AI{}/AI{}-1,"")".format(empty_row, one_year_ago)
Wish_list.cell(row = empty_row, column = col).number_format = percent_format
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)
#---------------------------------------------------------------------------------------------------------------
#column AN
col = col + 1
added_to_named_query = """
SELECT count(output.SKU)
                    from
                        (
                        select
                            wish_list_sku.sku SKU,
                             wish_list.id WL,
                             users.user_id
                        from
                            wish_list_sku,
                            wish_list,
                            users
                        where 1=1
                        and TO_CHAR(wish_list_sku.added,'YYYYMMDD') >= '{0}'
                        and TO_CHAR(wish_list_sku.added,'YYYYMMDD') <= '{1}'
                        and wish_list_sku.wish_list_id = wish_list.id
                        and wish_list.name !~* '(.*items i @heart@.*)'
                        and wish_list.user_id = users.user_id
                         and users.email !~* '(.*uncommongoods.*)'
                         and users.email !~* '(.*ugoods.*)'
                         and users.email !~* '(.*okeweka.*)'
                         and users.email !~* '(.*somethingsilly.*)'
                        group by wish_list_sku.sku, wish_list.id, users.user_id) output
""".format(__LAST_MONDAY__, __LAST_SUNDAY__)

added_to_named_df = Query('ugpostgres', added_to_named_query)
added_to_named = added_to_named_df.iloc[0] [0]

print('\n14. Skus Added to Named: ' + str(int(added_to_named)))

(Wish_list.cell(row = empty_row, column = col)).value = int(added_to_named)
Wish_list.cell(row = empty_row, column = col).number_format = '#,##0 '
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)
#---------------------------------------------------------------------------------------------------------------
#column space for calculation column AO (Unnamed Calculation)
col = col + 1
(Wish_list.cell(row = empty_row, column = col)).value = "=IFERROR(AN{}/AN{}-1,\"\")".format(empty_row, one_year_ago)
Wish_list.cell(row = empty_row, column = col).number_format = percent_format
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)
#---------------------------------------------------------------------------------------------------------------
#column AP
col = col + 1
added_to_loves_query = """SELECT count(ouput.SKU)
                    from
                        (
                        select
                            wish_list_sku.sku SKU,
                             wish_list.id WL,
                             users.user_id
                        from
                            wish_list_sku,
                            wish_list,
                            users
                        where 1=1
                            and TO_CHAR(wish_list_sku.added,'YYYYMMDD') >= '{0}'
                            and TO_CHAR(wish_list_sku.added,'YYYYMMDD') <= '{1}'
                            and wish_list_sku.wish_list_id = wish_list.id
                            and wish_list.name ~* '(.*items i @heart@.*)'
                            and wish_list.user_id = users.user_id
                             and users.email !~* '(.*uncommongoods.*)'
                             and users.email !~* '(.*ugoods.*)'
                             and users.email !~* '(.*okeweka.*)'
                             and users.email !~* '(.*somethingsilly.*)'
                            group by wish_list_sku.sku, wish_list.id, users.user_id) ouput
""".format(__LAST_MONDAY__,__LAST_SUNDAY__)
added_to_loves_df = Query('ugpostgres', added_to_loves_query)
added_to_loves = added_to_loves_df.iloc[0] [0]

print('\n15. Skus Added to Loves: ' + str(int(added_to_loves)))

(Wish_list.cell(row = empty_row, column = col)).value = int(added_to_loves)
Wish_list.cell(row = empty_row, column = col).number_format = '#,##0 '
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)
#---------------------------------------------------------------------------------------------------------------
#column space for calculation column AQ (Unnamed Calculation)
col = col + 1
(Wish_list.cell(row = empty_row, column = col)).value = "=IFERROR(AP{}/AP{}-1,\"\")".format(empty_row, one_year_ago)
Wish_list.cell(row = empty_row, column = col).number_format = percent_format
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)
#---------------------------------------------------------------------------------------------------------------
#Column AR
col = col + 1
skus_deleted_query = """SELECT count(SKU) DELETED_SKUS from
                (select wish_list_sku.sku SKU, wish_list.id, users.user_id
                from wish_list_sku, wish_list, users
                where TO_CHAR(wish_list_sku.deleted,'YYYYMMDD') >= '{0}'
                and TO_CHAR(wish_list_sku.deleted,'YYYYMMDD') <= '{1}'
                and wish_list_sku.wish_list_id = wish_list.id
                and wish_list.user_id = users.user_id
                 and users.email !~* '(.*uncommongoods.*)'
                 and users.email !~* '(.*ugoods.*)'
                 and users.email !~* '(.*okeweka.*)'
                 and users.email !~* '(.*somethingsilly.*)'
                group by wish_list_sku.sku, wish_list.id, users.user_id) output

""".format(__LAST_MONDAY__, __LAST_SUNDAY__)
skus_deleted_df = Query('ugpostgres', skus_deleted_query)
skus_deleted= skus_deleted_df.iloc[0][0]


print('\n16. Skus Deleted: ' + str(int(skus_deleted)))

(Wish_list.cell(row = empty_row, column = col)).value = int(skus_deleted)
Wish_list.cell(row = empty_row, column = col).number_format = '#,##0 '
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)
#---------------------------------------------------------------------------------------------------------------
#column space for calculation column AS (skus / customer)
col = col + 1

(Wish_list.cell(row= empty_row, column = col)).value = "=IFERROR(AI{}/E{},\"\")".format(empty_row, empty_row)
Wish_list.cell(row = empty_row, column= col).number_format = '#.## '
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)
#---------------------------------------------------------------------------------------------------------------
#column AT tot rev
col = col + 1

total_revenue_query ="""SELECT sum(output.REVENUE) Revenue,
                         count(distinct ORDER_ID) Orders,
                         sum(output.SKUS) Skus
                from
                    (
                    select
                            order_sku.order_sku_id ORDER_SKU,
                            min(orders.order_id) ORDER_ID,
                            min(order_sku.quantity) SKUS,
                            min(order_sku.quantity * order_sku.price) REVENUE
                    from
                            order_sku,
                            orders,
                            wish_list_sku,
                            wish_list,
                            users,
                            shipment

                    where 1=1
                    and order_sku.sku = wish_list_sku.sku
                    and shipment.order_id = orders.order_id
                    and shipment.authorized is not null
                    and shipment.is_cancelled = 0
                    and order_sku.is_cancelled = 0
                    and order_sku.order_id = orders.order_id
                    and orders.email = users.email
                    and wish_list.user_id = users.user_id
                    and wish_list_sku.wish_list_id = wish_list.id
                    and orders.date_created >= wish_list_sku.added
                    and TO_CHAR(orders.date_created,'YYYYMMDD') >= '{0}'
                    and TO_CHAR(orders.date_created,'YYYYMMDD') <= '{1}'
                    group by order_sku.order_sku_id
                    order by order_sku_id desc) output
""".format(__LAST_MONDAY__, __LAST_SUNDAY__)
total_revenue_df = Query('ugpostgres', total_revenue_query)
total_rev = total_revenue_df.iloc[0][0]


print('\n17. Total Rev: ' + str(int(total_rev)))

(Wish_list.cell(row = empty_row, column = col)).value = int(total_rev)
Wish_list.cell(row = empty_row, column = col).number_format = '$#,##0 '
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)
#---------------------------------------------------------------------------------------------------------------
#column space for calculation column AU (revenue growth calulation)
col = col + 1
(Wish_list.cell(row = empty_row, column = col)).value = "=IFERROR(AT{}/AT{}-1,\"\")".format(empty_row, one_year_ago)
Wish_list.cell(row = empty_row, column = col).number_format = percent_format
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)
#---------------------------------------------------------------------------------------------------------------
#Column AV
col = col + 1
total_revenue_query ="""SELECT sum(output.REVENUE) Revenue,
                         count(distinct ORDER_ID) Orders,
                         sum(output.SKUS) Skus
                from
                    (
                    select
                            order_sku.order_sku_id ORDER_SKU,
                            min(orders.order_id) ORDER_ID,
                            min(order_sku.quantity) SKUS,
                            min(order_sku.quantity * order_sku.price) REVENUE
                    from
                            order_sku,
                            orders,
                            wish_list_sku,
                            wish_list,
                            users,
                            shipment

                    where 1=1
                    and order_sku.sku = wish_list_sku.sku
                    and shipment.order_id = orders.order_id
                    and shipment.authorized is not null
                    and shipment.is_cancelled = 0
                    and order_sku.is_cancelled = 0
                    and order_sku.order_id = orders.order_id
                    and orders.email = users.email
                    and wish_list.user_id = users.user_id
                    and wish_list_sku.wish_list_id = wish_list.id
                    and orders.date_created >= wish_list_sku.added
                    and TO_CHAR(orders.date_created,'YYYYMMDD') >= '{0}'
                    and TO_CHAR(orders.date_created,'YYYYMMDD') <= '{1}'
                    group by order_sku.order_sku_id
                    order by order_sku_id desc) output
""".format(__LAST_MONDAY__, __LAST_SUNDAY__)
total_revenue_df = Query('ugpostgres', total_revenue_query)
total_sku = total_revenue_df.iloc[0][2]




print('\n18. Total Skus: ' + str(int(total_sku)))

(Wish_list.cell(row = empty_row, column = col)).value = int(total_sku)
Wish_list.cell(row = empty_row, column = col).number_format = '#,##0'
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)
# #---------------------------------------------------------------------------------------------------------------
#column space for calculation column AW (sku conv)
col = col + 1

(Wish_list.cell(row = empty_row, column = col)).value = "=IFERROR(AV224/(AI224+AI223),"")".format(empty_row, empty_row, (empty_row-1))
Wish_list.cell(row = empty_row, column= col).number_format = percent_format
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)
#---------------------------------------------------------------------------------------------------------------
#Column AX
col = col + 1

total_rev_sku_count_query = """SELECT sum(output.PRICE) REVENUE, count(output.WLSKUID) SKUS
                from
                    (
                        select wish_list_sku.id WLSKUID,
                             sku.price PRICE
                        from wish_list_sku, wish_list, users, sku
                        where wish_list_sku.added_to_cart is not null
                        and wish_list_sku.wish_list_id = wish_list.id
                        and wish_list.user_id = users.user_id
                        and wish_list_sku.sku = sku.sku
                        and TO_CHAR(wish_list_sku.added_to_cart,'YYYYMMDD') >= '{0}'
                        and TO_CHAR(wish_list_sku.added_to_cart,'YYYYMMDD') <= '{1}'
                         and users.email !~* '(.*uncommongoods.*)'
                         and users.email !~* '(.*ugoods.*)'
                         and users.email !~* '(.*okeweka.*)'
                         and users.email !~* '(.*somethingsilly.*)'
                        group by wish_list_sku.id, sku.price) output

""".format(__LAST_MONDAY__, __LAST_SUNDAY__)
total_rev_sku_count_df = Query('ugpostgres', total_rev_sku_count_query)
rev_moved = total_rev_sku_count_df.iloc[0][0]


print('\n19. Rev Moved: ' + str(int(rev_moved)))

(Wish_list.cell(row = empty_row, column = col)).value = int(rev_moved)
Wish_list.cell(row = empty_row, column = col).number_format = '#,##0 '
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)


#---------------------------------------------------------------------------------------------------------------------
#Column AY
col = col + 1
sku_moved= total_rev_sku_count_df.iloc[0][1]

print('\n20. sku moved: ' + str(int(sku_moved)))

(Wish_list.cell(row = empty_row, column = col)).value = int(sku_moved)
Wish_list.cell(row = empty_row, column = col).number_format = '#,##0 '
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)

#---------------------------------------------------------------------------------------------------------------
#Column AZ
col = col + 1

total_rev_sku_count_with_match_query = """SELECT sum(PRICE) REVENUE, count(SKU_IDS) SKUS from
                (select wish_list_sku.id SKU_IDS, sku.price PRICE, order_sku.order_sku_id
                from wish_list_sku, wish_list, users, sku, order_sku, orders, shipment
                where wish_list_sku.added_to_cart is not null
                and wish_list_sku.wish_list_id = wish_list.id
                and wish_list.user_id = users.user_id
                and users.email = orders.email
                and orders.order_id = order_sku.order_id
                and order_sku.sku = sku.sku
                and sku.sku = wish_list_sku.sku
                and shipment.order_id = orders.order_id
                and shipment.authorized is not null
                and shipment.is_cancelled = 0
                and order_sku.is_cancelled = 0
                 and users.email !~* '(.*uncommongoods.*)'
                 and users.email !~* '(.*ugoods.*)'
                 and users.email !~* '(.*okeweka.*)'
                 and users.email !~* '(.*somethingsilly.*)'
                and TO_CHAR(wish_list_sku.added_to_cart,'YYYYMMDD') >= '{0}'
                and TO_CHAR(wish_list_sku.added_to_cart,'YYYYMMDD') <= '{1}'
                group by wish_list_sku.id, sku.price, order_sku.order_sku_id) output

""".format(__LAST_MONDAY__, __LAST_SUNDAY__)
total_rev_sku_count_with_match_df = Query('ugpostgres',total_rev_sku_count_with_match_query)
rev_moved_with_order = total_rev_sku_count_with_match_df.iloc[0][0]


print('\n20. Rev Moved With Order: ' + str(int(rev_moved_with_order)))

(Wish_list.cell(row = empty_row, column = col)).value = int(rev_moved_with_order)
Wish_list.cell(row = empty_row, column = col).number_format = '#,##0 '
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)
#---------------------------------------------------------------------------------------------------------------
#Column BA
col = col + 1
sku_moved_with_order = total_rev_sku_count_with_match_df.iloc[0][1]


print('\n21. Rev Moved With Order: ' + str(int(sku_moved_with_order)))

(Wish_list.cell(row = empty_row, column = col)).value = int(sku_moved_with_order)
Wish_list.cell(row = empty_row, column = col).number_format = '#,##0 '
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)
#---------------------------------------------------------------------------------------------------------------
#Column BB
col = col + 1

orders_total_revenue_query = """
SELECT sum(output.REVENUE) Revenue,
                         count(distinct ORDER_ID) Orders,
                         sum(output.SKUS) Skus
                from
                    (
                    select
                            order_sku.order_sku_id ORDER_SKU,
                            min(orders.order_id) ORDER_ID,
                            min(order_sku.quantity) SKUS,
                            min(order_sku.quantity * order_sku.price) REVENUE
                    from
                            order_sku,
                            orders,
                            wish_list_sku,
                            wish_list,
                            users,
                            shipment

                    where 1=1
                    and order_sku.sku = wish_list_sku.sku
                    and shipment.order_id = orders.order_id
                    and shipment.authorized is not null
                    and shipment.is_cancelled = 0
                    and order_sku.is_cancelled = 0
                    and order_sku.order_id = orders.order_id
                    and orders.email = users.email
                    and wish_list.user_id = users.user_id
                    and wish_list_sku.wish_list_id = wish_list.id
                    and orders.date_created >= wish_list_sku.added
                    and TO_CHAR(orders.date_created,'YYYYMMDD') >= '{0}'
                    and TO_CHAR(orders.date_created,'YYYYMMDD') <= '{1}'
                    group by order_sku.order_sku_id
                    order by order_sku_id desc) output
""".format(__LAST_MONDAY__, __LAST_SUNDAY__)
orders_total_revenue_df = Query('ugpostgres', orders_total_revenue_query)
orders_total_revenue = orders_total_revenue_df.iloc[0] [1]


print('\n22. Total orders ' + str(int(orders_total_revenue)))

(Wish_list.cell(row = empty_row, column = col)).value = int(orders_total_revenue)
Wish_list.cell(row = empty_row, column = col).number_format = '#,##0 '
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)
# #---------------------------------------------------------------------------------------------------------------
#Column BC
col = col + 1
AOV_query = """
SELECT round(sum(order_sku.price * order_sku.quantity) / count(distinct WL_ORDERS.order_id)) AOV,
                sum(order_sku.price * order_sku.quantity) REV
                from order_sku, (select distinct orders.order_id
                from order_sku, orders, wish_list_sku, wish_list, users
                where order_sku.sku = wish_list_sku.sku
                and order_sku.order_id = orders.order_id
                and orders.email = users.email
                and wish_list.user_id = users.user_id
                and wish_list_sku.wish_list_id = wish_list.id
                and orders.date_created >= wish_list_sku.added
                and TO_CHAR(orders.date_created,'YYYYMMDD') >= '{0}'
                and TO_CHAR(orders.date_created,'YYYYMMDD') <= '{1}') WL_ORDERS
                where order_sku.order_id = WL_ORDERS.order_id
""".format(__LAST_MONDAY__, __LAST_SUNDAY__)
AOV_query_df = Query('ugpostgres', AOV_query)
AOV = AOV_query_df.iloc[0][0]


print('\n23. AOV: ' + str(int(AOV)))

(Wish_list.cell(row = empty_row, column = col)).value = int(AOV)
Wish_list.cell(row = empty_row, column = col).number_format = '$#,##0 '
Wish_list.cell(row = empty_row, column = col).font = Font(size=12)
#---------------------------------------------------------------------------------------------------------------
#---------------------------------------------------------------------------------------------------------------
#---------------------------------------------------------------------------------------------------------------#---------------------------------------------------------------------------------------------------------------
#---------------------------------------------------------------------------------------------------------------
#---------------------------------------------------------------------------------------------------------------
#---------------------------------------------------------------------------------------------------------------



























#RuxWishList.save('/Users/gconnolly/Documents/projects/RUX/RUX_automated.xlsx')
RuxWishList.save('/Users/gconnolly/Documents/projects/RUX/test.xlsx')
