#NOTE: All numbers defining 'spikes' and 'driving product catalog numbers' in this code are arbitary. They have yet to be optimized
#Changing the numbers to be lower, will allow more noise created in the graphs, excel tables, and text files
#Changing the numbers to be higher, will miss important information
#These numbers should be changed to fit the user's need and optimized per the scenario this code is used
#NOTE: All of the headers of the data excel file must be consistent with the names of the predicted columns in this code
#If the name changes, it must be adjusted here
#For example: if the excel file data column for product minor is called P. Minor instead of Product Minor, this code will not work until it is updated


#Useful Generators
def star_generator(arg):
    """Creates a string of stars to be used with formatting text files
    
    Arguments:
        arg {integer} -- Number of stars added to string
    
    Yields:
        {string} -- star
    """
    for x in range(arg):
        x = x
        yield '*'

def dash_generator(arg):
    """Creates a string of dashes to be used with formating text files
    
    Arguments:
        arg {integer} -- Number of dashes added to string
    
    Yields:
        {string} -- dash
    """
    for x in range(arg):
        x = x
        yield '-'

#Functions used to decipher details and determine potential similarities between complaints
def detail_hospital(data,list_current_month,month_names,month,star,dash,product_index,trending_index,minor_index,name_file_text):
    """Provides a tool for analyzing the data by sorting for reoccurrent complaintant facilities
    
    Arguments:
        data {dictionary} -- Complaint data pulled from user's file
        list_current_month {list} -- Complaint data filtered by Product Major Group and Month. Note each complaint's details is a list within the list
        month_names {dictionary} -- Months of the year
        month {string} -- Most recent month from the complaint data
        star {string} -- Line of stars for aesthetics
        dash {string} -- Line of dashes for aesthetics
        product_index {integer} -- Index for product catalog number per complaint list
        trending_index {integer} -- Index for trending device code per complaint list
        minor_index {integer} -- Index for product minor group per complaint list
        name_file_text {string} -- File location within the users directory 
    """
    #Find the index for where Complainant Facility is located within data
    index = -1
    for x in data:
        index += 1
        if x == 'Complainant Facility':
            cf_index = index
    #Create a new list for complaintant facility for each relevant complaint
    hospital = []
    for x in list_current_month:
        hospital.append(x[cf_index])
    hospital_set = [{a for a in hospital}] #Reduce the number of complaintant facilities by removing dublicates
    #Sort through the complaints and give information on the complaintant facility, trending device codes affected, product minor groups affected, and product codes affected
    for x in hospital_set:
        for y in x:
            if hospital.count(y) > 3: #Look for trend in the hosptials reporting complaints. The nuber defines the minimum complaints to alert the user of a possible trend
                #If there is a potential hospital trend, format the text file
                file_text = open(name_file_text,'a')
                file_text.write('\n\n' + star + '\n' + star)
                file_text.close()
                break
    for x in hospital_set:
        for y in x:
            if hospital.count(y) > 3: #Look for trend in the hosptials reporting complaints. The nuber defines the minimum complaints to alert the user of a possible trend
                #Add the details of the potential hospital trend to the text file
                file_text = open(name_file_text,'a')
                file_text.write('\n\n\n\tWARNING! COMPLAINTANT FACILITY DUMP')
                file_text.write('\n%s had %d complaints in the month of %s' % (y,hospital.count(y),month_names[str(month)]))
                #Include the trending device codes, product minor groups, and product catalog numbers in the details
                file_text.write('\nAffected Trending Device Codes: ')
                hospital_trending = list(filter(lambda z: z[cf_index] == y,list_current_month))
                trending_per_h = []
                for z in hospital_trending:
                    trending_per_h.append(z[trending_index])
                trending_per_h_set = [{a for a in trending_per_h}]
                for z in trending_per_h_set:
                    for q in z:
                        file_text.write('\n\t-%s with %d complaints' % (q,trending_per_h.count(q)))
                file_text.write('\nAffected Product Minors: ') 
                minor_per_h = []
                for z in hospital_trending:
                    minor_per_h.append(z[minor_index])
                minor_per_h_set = [{a for a in minor_per_h}]
                for z in minor_per_h_set:
                    for q in z:
                        file_text.write('\n\t%s:' % q)
                        product_per_h = []
                        for t in hospital_trending:
                            if t[minor_index] == q:
                                product_per_h.append(t[product_index])
                        product_per_h_set = [{a for a in product_per_h}]
                        for t in product_per_h_set:
                            for u in t:
                                file_text.write('\n\t-%r had %d complaints' % (u,product_per_h.count(u)))
                file_text.write('\n\n' + dash + '\n')
                file_text.close()
    for x in hospital_set:
        for y in x:
            if hospital.count(y) > 3: #Look for trend in the hosptials reporting complaints. The nuber defines the minimum complaints to alert the user of a possible trend
                #If there is a potential hospital trend, format the text file (again)
                file_text = open(name_file_text,'a')
                file_text.write('\n\n' + star + '\n' + star)
                file_text.close()
                break
    return

def detail_lot(data,list_current_month,month_names,month,star,dash,product_index,trending_index,minor_index,name_file_text):
    """Provides a tool for analyzing the data by sorting for reoccurrent lots
    
    Arguments:
        data {dictionary} -- Complaint data pulled from user's file
        list_current_month {list} -- Complaint data filtered by Product Major Group and Month. Note each complaint's details is a list within the list
        month_names {dictionary} -- Months of the year
        month {string} -- Most recent month from the complaint data
        star {string} -- Line of stars for aesthetics
        dash {string} -- Line of dashes for aesthetics
        product_index {integer} -- Index for product catalog number per complaint list
        trending_index {integer} -- Index for trending device code per complaint list
        minor_index {integer} -- Index for product minor group per complaint list
        name_file_text {string} -- File location within the users directory 
    """
    #Find the indices for where lot information is located within data
    index = -1
    for x in data:
        index += 1
        if x == 'Corporate Lot No':
            cln_index = index
    index = -1
    for x in data:
        index += 1
        if x == 'Manufacturing Lot No':
            mln_index = index
    #When the complaint details do not have lot information, replace it with an uniform 'UNK' to make tracking it easier
    for x in list_current_month:
        try:
            if x[cln_index].upper() == 'UNK':
                x[cln_index] = 'UNK'
            elif x[cln_index].upper() == 'UNKNOWN':
                x[cln_index] = 'UNK'
            elif x[cln_index].upper() == 'NA':
                x[cln_index] = 'UNK'
        except:
            pass
        if x[cln_index] == 0:
            x[cln_index] = 'UNK'
        elif x[cln_index] == 'n/a' or x[cln_index] == 'N/A':
            x[cln_index] ='UNK'

        try:
            if x[mln_index].upper() == 'UNK':
                x[mln_index] = 'UNK'
            elif x[mln_index].upper() == 'UNKNOWN':
                x[mln_index] = 'UNK'
            elif x[mln_index].upper() == 'NA':
                x[mln_index] = 'UNK'
        except:
            pass
        if x[mln_index] == 0:
            x[mln_index] = 'UNK'
        elif x[mln_index] == 'n/a' or x[mln_index] == 'N/A':
                x[mln_index] ='UNK'
    #Create a new list for lot information for each relevant complaint
    corporate = []
    for x in list_current_month:
        corporate.append(x[cln_index])
    corporate_set = [{a for a in corporate}] #Reduce the lot information by removing dublicates
    for x in corporate_set:
        if 'UNK' in x:
            x.remove('UNK') #Remove the unkown lots so UNK does not appear as a trend
    #Sort through the complaints and give information on the lot, trending device codes affected, product minor groups affected, and product catalog numbers affected
    start_print = 0
    for x in corporate_set:
        for y in x:
            if corporate.count(y) > 3: #Look for trend in the lots. The nuber defines the minimum complaints to alert the user of a possible trend
                #If there is a potential lot trend, format the text file
                file_text = open(name_file_text,'a')
                file_text.write('\n\n' + star + '\n' + star)
                start_print = 1
                file_text.close()
                break
    for x in corporate_set:
        for y in x:
            if corporate.count(y) > 3: #Look for trend in the lots. The nuber defines the minimum complaints to alert the user of a possible trend
                #Add the details of the potential lot trend to the text file
                file_text = open(name_file_text,'a')
                file_text.write('\n\n\tWARNING! CORPORATE LOT NUMBER')
                file_text.write('\n%r had %d complaints in the month of %s' % (y,corporate.count(y),month_names[str(month)]))
                #Include the trending device codes, product minor groups, and product catalog numbers in the details
                file_text.write('\nAffected Trending Device Codes: ')
                corporate_trending = list(filter(lambda z: z[cln_index] == y,list_current_month))
                trending_per_c = []
                for z in corporate_trending:
                    trending_per_c.append(z[trending_index])
                trending_per_c_set = [{a for a in trending_per_c}]
                for z in trending_per_c_set:
                    for q in z:
                        file_text.write('\n\t-%s with %d complaints' % (q,trending_per_c.count(q)))
                file_text.write('\nAffected Product Minors: ')
                minor_per_c = []
                for z in corporate_trending:
                    minor_per_c.append(z[minor_index])
                minor_per_c_set = [{a for a in minor_per_c}]
                for z in minor_per_c_set:
                    for q in z:
                        file_text.write('\n\t%s:' % q)
                        product_per_c = []
                        for t in corporate_trending:
                            if t[minor_index] == q:
                                product_per_c.append(t[product_index])
                        product_per_c_set = [{a for a in product_per_c}]
                        for t in product_per_c_set:
                            for u in t:
                                file_text.write('\n\t-%r had %d complaints' % (u,product_per_c.count(u)))
                file_text.write('\n\n' + dash + '\n')
                file_text.close()
    manufacture = list(filter(lambda x: x[cln_index] == 'UNK',list_current_month))
    man_lot = []
    for x in manufacture:
        man_lot.append(x[mln_index])
    man_lot_set = [{a for a in man_lot}]
    for x in man_lot_set:
        if 'UNK' in x:
            x.remove('UNK') #Remove the unkown lots so UNK does not appear as a trend
    if start_print == 0: #If the text file has not been formatted yet, see if it needs to be
        for x in man_lot_set:
            for y in x:
                if man_lot.count(y) > 3: #Look for trend in the lots. The nuber defines the minimum complaints to alert the user of a possible trend
                    #If there is a potential lot trend, format the text file
                    file_text = open(name_file_text,'a')
                    file_text.write('\n\n' + star + '\n' + star)
                    file_text.close()
                    break
    for x in man_lot_set:
        for y in x:
            if man_lot.count(y) > 3: #Look for trend in the lots. The nuber defines the minimum complaints to alert the user of a possible trend
                #Add the details of the potential lot trend to the text file
                file_text = open(name_file_text,'a')
                file_text.write('\n\n\tWARNING! MANUFACTURING LOT NUMBER')
                file_text.write('\n%r had %d complaints in the month of %s' % (y,man_lot.count(y),month_names[str(month)]))
                #Include the trending device codes, product minor groups, and product catalog numbers in the details
                file_text.write('\nAffected Trending Device Codes: ')
                man_trending = list(filter(lambda z: z[mln_index] == y,list_current_month))
                trending_per_m = []
                for z in man_trending:
                    trending_per_m.append(z[trending_index])
                trending_per_m_set = [{a for a in trending_per_m}]
                for z in trending_per_m_set:
                    for q in z:
                        file_text.write('\n\t-%s with %d complaints' % (q,trending_per_m.count(q)))
                file_text.write('\nAffected Product Minors: ')
                minor_per_m = []
                for z in man_trending:
                    minor_per_m.append(z[minor_index])
                minor_per_m_set = [{a for a in minor_per_m}]
                for z in minor_per_m_set:
                    for q in z:
                        file_text.write('\n\t%s:' % q)
                        product_per_m = []
                        for t in man_trending:
                            if t[minor_index] == q:
                                product_per_m.append(t[product_index])
                        product_per_m_set = [{a for a in product_per_m}]
                        for t in product_per_m_set:
                            for u in t:
                                file_text.write('\n\t-%r had %d complaints' % (u,product_per_m.count(u)))
                file_text.write('\n\n' + dash + '\n')
                file_text.close()
    for x in corporate_set:
        for y in x:
            if corporate.count(y) > 3: #Look for trend in the lots. The nuber defines the minimum complaints to alert the user of a possible trend
                #If there is a potential lot trend, format the text file
                file_text = open(name_file_text,'a')
                file_text.write('\n\n' + star + '\n' + star)
                file_text.close()
                break
    if start_print == 0: #If the text file has not been formatted yet, see if it needs to be
        for x in man_lot_set:
            for y in x:
                if man_lot.count(y) > 3: #Look for trend in the lots. The nuber defines the minimum complaints to alert the user of a possible trend
                    #If there is a potential lot trend, format the text file
                    file_text = open(name_file_text,'a')
                    file_text.write('\n\n' + star + '\n' + star)
                    file_text.close()
                    break
    return 

def detail_quantity(data,list_current_month,month_names,month,star,dash,product_index,trending_index,minor_index,name_file_text):
    """Provides a tool for analyzing the data by sorting for high quantities affected by complaints
    
    Arguments:
        data {dictionary} -- Complaint data pulled from user's file
        list_current_month {list} -- Complaint data filtered by Product Major Group and Month. Note each complaint's details is a list within the list
        month_names {dictionary} -- Months of the year
        month {string} -- Most recent month from the complaint data
        star {string} -- Line of stars for aesthetics
        dash {string} -- Line of dashes for aesthetics
        product_index {integer} -- Index for product catalog number per complaint list
        trending_index {integer} -- Index for trending device code per complaint list
        minor_index {integer} -- Index for product minor group per complaint list
        name_file_text {string} -- File location within the users directory 
    """
    #Find the index for where Quantity Affected is located within data
    index = -1
    for x in data:
        index += 1
        if x == 'Quantity Affected':
            qa_index = index
    index = -1
    for x in data:
        index += 1
        if x == 'Complainant Facility':
            cf_index = index
    #If the quantity affected is unknown, assume 1 (conservative approach)
    quantity_affected = []
    for x in list_current_month:
        try:
            x[qa_index] = int(x[qa_index])
        except:
            x[qa_index] = 1
        quantity_affected.append(x[qa_index])
    if len(quantity_affected) == 0: #If there is not any complaints for this product, set the max to zero so the program does not crash
        max_qa = 0
    else: 
        max_qa = max(quantity_affected) 
    if max_qa > 4: #The number sets the minimum needed to alert the user if there is a potential high number of quantities affected reported for a complaint
        #Add the details to the text file
        file_text = open(name_file_text,'a')
        file_text.write('\n\n' + star + '\n' + star)
        file_text.write('\n\n\tWARNING! QUANTITY AFFECTED')
        for x in list_current_month:
            if x[qa_index] > 4: #The number sets the minimum needed to alert the user if there is a potential high number of quantities affected reported for a complaint
                file_text.write('\n%d quantity affected for a complaint submitted by %s for Trending Device Code %s for product %r (%s)' % (x[qa_index],x[cf_index],x[trending_index],x[product_index],x[minor_index])) ###***
        file_text.write('\n\n' + star + '\n' + star)
        file_text.close()
    return 

#Coding below is used to sort the data and create excel sheets
def mainfunc():
    '''Mainfunc is the primary function for the program. All other functions are eventually used within this function.
    Data will be gathered from the current directory. An excel summary will be made for each product major with hyperlinked text files
    containing additional information. The most recent month will be identified in the function and be used as the focal point. 
    Each summary will explore spikes in complaints for trending device codes, product minor groups, and individual product catalog numbers.
    No parameters are needed to start this function and no returns are yielded by this function
    '''
    #Useful imports
    import pandas as pd
    import numpy as np
    import matplotlib.pyplot as plt
    import os
    import shutil
    import openpyxl
    from openpyxl import load_workbook
    from openpyxl import Workbook
    from openpyxl.styles import Font, colors
    #Use generators to create strings for formatting text files
    star = ''
    for x in list(star_generator(150)):
        star += x
    
    dash = ''
    for x in list(dash_generator(150)):
        dash += x

    #Find the current working directory and open the data named 'complaint_detail_report.csv' within it
    rx = os.getcwd()
    data = pd.read_csv((rx + '\\complaint_detail_report.csv'),engine='python')
    data = data.dropna(how='all') # Drop all N/A values (gets rid of blank rows at the bottom of the import)
    #Create folders for the text files
    rxr = rx + '\\' + 'Text File Dump'
    rxred = rx + '\\' + 'Event Descriptions Dump'
    #Look to see if there is text files and/or excel summaries from last month. If so, delete them so they can be replaced
    try:
        shutil.rmtree(rxr)
    except:
        print('Note: Text File Dump folder was not detected.')
    try:
        shutil.rmtree(rxred)
    except:
        print('Note: Event Descriptions Dump folder was not detected.')
    directory = [a for a in rx.split('\\')] #The structure has the excel summaries back a level (from the folder containing the program) and inside a separate folder
    new_directory = ''
    for x in directory[0:(len(directory)-1)]:
        new_directory += (x + '\\')
    rru = new_directory + 'Product Major Complaint Summaries\\'
    try:
        shutil.rmtree(rru)
    except:
        print('Note: Product Major Complaint Summaries folder was not detected')
    #Create new directories for the text files and excel summaries 
    os.mkdir(rru) #Excel summaries 
    os.mkdir(rxred) #Event descriptions (text file)

    product_major = data['Product Major Group']
    product_major_set = [{a for a in product_major}] #Create a set of all the product majors with complaints
    #Note the below attempts to fill N/A values is imperfect. Sometimes, N/A values slip through. Functions above address those cases.
    data['Corporate Lot No'].fillna(0, inplace=True) #Replace all numpy.nan (not numbers or letters, comes from a cell with NA) with 0
    data['Manufacturing Lot No'].fillna(0, inplace=True) #Replace all numpy.nan (not numbers or letters, comes from a cell with NA) with 0
    #Find the latest month from the complaint data
    month_names = {'1':'January','2':'February','3':'March','4':'April','5':'May','6':'June','7':'July','8':'August','9':'September','10':'October','11':'November','12':'December'}
    time = data['Date Opened'] 
    date = time[(len(time)-1)] 
    if '-' in date: #Sometimes there are inconsistencies with the dates. Splice the dates by their dividers
        date = [a for a in time[(len(time) - 1)].split('-')]
        month = date[1]
    else:
        date = [a for a in time[(len(time) - 1)].split('/')]
        month = date[0]

    os.mkdir(rxr) #Text files (additional details)
    #For each product major in the complaint data, create a new folder and excel summary 
    for o in product_major_set:
        for e in o:
            preference = e #Product major 
            rr = rru + preference 
            os.mkdir(rr) #Individual excel summary
            #Create a text file for additional details
            name_file_text = rxr + '\\' + preference + '.txt'
            file_text = open(name_file_text,'w')
            #Header of text file
            file_text.write('This is a text file with \'WARNING\'s for: \n\t1) Complaintant Facility Dumps \n\t2) Recurring Lot Numbers (Coporate and Manufacturer) \n\t3) High Quantity Affected Units per Complaints.\n\n' + star + '\n' + dash + '\n' + star + '\n\n\n')
            file_text.close()

            #Find all the important indices
            index = -1
            for x in data:
                index += 1
                if x == 'Date Opened':
                    date_index = index
                elif x == 'Trending Device Codes':
                    trending_index = index
                elif x == 'Product Minor Group':
                    minor_index = index
                elif x == 'Product Catalog No.':
                    product_index = index
            #Create a list of indices to indentify each complaint related to the product major
            list_indices = []
            column = data['Product Major Group']
            index = -1
            for x in column:
                index += 1
                if x == preference:
                    list_indices.append(index)
            #Take the data from the dictionary, compile complaint details into a list and add each list to a master list
            list_major = []
            for x in list_indices: #Only complaint data for the current product major will be extracted
                row = []
                for y in data:
                    z = data[y]
                    row.append(z[x])
                list_major.append(row)
            #Using list major, segregate the data by months. A new list for each month will be made for the past 6 months and added to master list
            past_months = int(month)
            list_multiple_embedded = []
            all_months = [] #Keeping track of the months being analyzed
            all_months.append(int(month))
            for x in range(7):
                list_major_past = []
                for y in list_major:
                    list_date = y[date_index]
                    if '-' in list_date: #Sometimes there are inconsistencies with the dates. Splice the dates by their dividers
                        list_date = [a for a in y[date_index].split('-')]
                        if '0' in list_date[1] and list_date[1] != '10':
                            list_date[1] = list_date[1].replace('0','') #Change 01 to 1, but make sure its not 10 that's being changed
                        if list_date[1] == str(past_months):
                            list_major_past.append(y)
                    else:
                        list_date = [a for a in y[date_index].split('/')]
                        if list_date[0] == str(past_months):
                            list_major_past.append(y)
                list_multiple_embedded.append(list_major_past)
                if past_months - 1 == 0:
                    past_months = 12 #Before January is December
                else:
                    past_months -= 1
                all_months.append(past_months)
            #Create new lists for tracking complaints per trending device codes, product minor groups, and product catalog numbers
            minors_previous = []
            trending_previous = []
            product_previous = []
            index = -1
            for x in list_multiple_embedded:
                index += 1
                for y in x:
                    for z in y:
                        if y.index(z) == minor_index:
                            minors_previous.append(z)
                        elif y.index(z) == trending_index:
                            trending_previous.append(z)
                        elif y.index(z) == product_index:
                            product_previous.append(z)
            #Create sets of the new lists to remove dublicates
            trending_set_previous = [{a for a in trending_previous}]
            minors_set_previous = [{a for a in minors_previous}]
            product_set_previous = [{a for a in product_previous}]
            #Create a list that will have embedded lists with information per month of the amount of complaints per trending device code, product minor group, and product catalog
            master_data = []
            for x in list_multiple_embedded:
                one_data = [] #Used for compliation of tracking information
                trending_nonset = [] #Making nonsets of the set data. Having set data embedded was difficult to use
                trending_count = []
                for y in trending_set_previous:
                    for z in y:
                        trending_indv = 0
                        trending_nonset.append(z)
                        for q in x:
                            if q[trending_index] == z:
                                trending_indv += 1
                        trending_count.append(trending_indv) #Keeping track of the number of complaints per trending device code in the month
                one_data.append(trending_nonset)
                one_data.append(trending_count)

                minors_nonset = [] #Making nonsets of the set data. Having set data embedded was difficult to use
                minors_count = []
                for y in minors_set_previous:
                    for z in y:
                        minors_indv = 0
                        minors_nonset.append(z)
                        for q in x:
                            if q[minor_index] == z:
                                minors_indv += 1
                        minors_count.append(minors_indv) #Keeping track of the number of complaints per trending device code in the month
                one_data.append(minors_nonset)
                one_data.append(minors_count)

                product_nonset = [] #Making nonsets of the set data. Having set data embedded was difficult to use
                product_count = []
                for y in product_set_previous:
                    for z in y:
                        product_indv = 0
                        product_nonset.append(z)
                        for q in x:
                            if q[product_index] == z:
                                product_indv += 1
                        product_count.append(product_indv) #Keeping track of the number of complaints per trending device code in the month
                one_data.append(product_nonset)
                one_data.append(product_count)
                master_data.append(one_data)
            #Create a dictionary to compile complaint spikes
            spike = {} #A dictionary is used because it is more compatible for pandas to important into excel files
            #Current month minus last month compared to a minimum increase in complaints defines a spike
            current_data = master_data[0]
            previous_data = master_data[1]
            #Find spikes for trending device codes 
            trending_dc_current = current_data[0]
            trending_current_data = current_data[1]
            trending_previous_data = previous_data[1]
            flag = []
            complaints_flag = []
            complaints_diff_flag = []
            index = -1 #Index will be used to manually index through the previous data while the current data is index through a for loop
            for x in trending_current_data:
                index += 1
                #The amount of complaints will influence the threshold for a spike
                if x < 10:
                    if x - trending_previous_data[index] > 1: #If the complaints are less than 10, an increase by 2 is a spike
                        flag.append(trending_dc_current[index]) #Track the trending device code with a spike
                        complaints_flag.append(x) #Track the amount of complaints
                        complaints_diff_flag.append(x - trending_previous_data[index]) #Track the magnitude of the spike
                elif x < 20:
                    if x - trending_previous_data[index] > 2: #If the complaints are less than 20, an increase by 3 is a spike
                        flag.append(trending_dc_current[index])
                        complaints_flag.append(x)
                        complaints_diff_flag.append(x - trending_previous_data[index])
                elif x < 40:
                    if x - trending_previous_data[index] > 3: #If the complaints are less than 40, an increase by 4 is a spike
                        flag.append(trending_dc_current[index])
                        complaints_flag.append(x)
                        complaints_diff_flag.append(x - trending_previous_data[index])
                else:
                    if (((x - trending_previous_data[index])/x)*100) >= 10: #If the complaints are 40 or greater, an increase by 10% is a spike
                        flag.append(trending_dc_current[index])
                        complaints_flag.append(x)
                        complaints_diff_flag.append(x - trending_previous_data[index])
            #Add the tracked information to the dictionary 'spike'
            spike['Trending Device Codes'] = flag
            spike['TDC %s Complaints' % month_names[month]] = complaints_flag
            spike['TDC Monthly Complaint Increase'] = complaints_diff_flag
            #Find spikes for product minors 
            minors_current = current_data[2]
            minors_current_data = current_data[3]
            minors_previous_data = previous_data[3]
            flag = []
            complaints_flag = []
            complaints_diff_flag = []
            index = -1 #Index will be used to manually index through the previous data while the current data is index through a for loop
            for x in minors_current_data:
                index += 1
                #The amount of complaints will influence the threshold for a spike
                if x < 10:
                    if x - minors_previous_data[index] > 1: #If the complaints are less than 10, an increase by 2 is a spike
                        flag.append(minors_current[index]) #Track the trending device code with a spike
                        complaints_flag.append(x) #Track the amount of complaints
                        complaints_diff_flag.append(x - minors_previous_data[index]) #Track the magnitude of the spike
                elif x < 20:
                    if x - minors_previous_data[index] > 2: #If the complaints are less than 20, an increase by 3 is a spike
                        flag.append(minors_current[index])
                        complaints_flag.append(x)
                        complaints_diff_flag.append(x - minors_previous_data[index])
                elif x < 40:
                    if x - minors_previous_data[index] > 3: #If the complaints are less than 40, an increase by 4 is a spike
                        flag.append(minors_current[index])
                        complaints_flag.append(x)
                        complaints_diff_flag.append(x - minors_previous_data[index])
                else:
                    if (((x - minors_previous_data[index])/x)*100) >= 10: #If the complaints are 40 or greater, an increase by 10% is a spike
                        flag.append(minors_current[index])
                        complaints_flag.append(x)
                        complaints_diff_flag.append(x - minors_previous_data[index])
            #Add the tracked information to the dictionary 'spike'
            spike['Product Minor Groups'] = flag
            spike['PMG %s Complaints' % month_names[month]] = complaints_flag
            spike['PMG Monthly Complaint Increase'] = complaints_diff_flag
            #Find spikes for product catalog numbers 
            product_current = current_data[4]
            product_current_data = current_data[5]
            product_previous_data = previous_data[5]
            flag = []
            complaints_flag = []
            complaints_diff_flag = []
            index = -1
            for x in product_current_data:
                index += 1
                #The amount of complaints will influence the threshold for a spike
                #Product catalog numbers are more specific than trending device codes or product minors. This causes the filter for less than 10 complaints to create too much noise.
                if x < 20:
                    if x - product_previous_data[index] > 2: #If the complaints are less than 20, an increase by 3 is a spike
                        flag.append(product_current[index])
                        complaints_flag.append(x)
                        complaints_diff_flag.append(x - product_previous_data[index])
                elif x < 40:
                    if x - product_previous_data[index] > 3: #If the complaints are less than 40, an increase by 4 is a spike
                        flag.append(product_current[index])
                        complaints_flag.append(x)
                        complaints_diff_flag.append(x - product_previous_data[index])
                else:
                    if (((x - product_previous_data[index])/x)*100) >= 10: #If the complaints are 40 or greater, an increase by 10% is a spike
                        flag.append(product_current[index])
                        complaints_flag.append(x)
                        complaints_diff_flag.append(x - product_previous_data[index])
            #Add the tracked information to the dictionary 'spike'
            spike['Product Catalog Numbers'] = flag
            spike['PCN %s Complaints' % month_names[month]] = complaints_flag
            spike['PCN Monthly Complaint Increase'] = complaints_diff_flag

            #Create a text file for event descriptions
            name_file_ed = rxred + '\\' + preference + '.txt'
            file_ed = open(name_file_ed,'w')
            file_ed.write('This is a text file with Event Descriptions segregated by flagged Trending Device Codes, Product Minors, and Product Catalog Numbers.\nIf the flagged Trending Device Code or Product Minor has driving Product Catalog Numbers, \nthe event descriptions are filtered to only show the driving Product Catalog Numbers WITHIN the Trending Device Code / Product Minor' + '\n' + star + '\n' + dash + '\n' + star + '\n\n')
            file_ed.close()
            #Find the index for event descriptions within the data
            index = -1
            for x in data:
                index += 1
                if x == 'Event Description':
                    ed_index = index

            #Extract the infromation from the first dictionary to create a new dictionary. This dictionary will be more specific and only include trending device codes
            #The event descriptions text file will be updated simultaneously as the new dictionary informaiton is organized
            trending_products = {} #New dictionary specific to trending device codes
            trending_flagged = spike['Trending Device Codes']
            #Create new lists for organizing data for the new dictionary
            product_trending_tdc = []
            product_trending_p =[]
            product_trending_comp = []
            product_trending_diff = []
            #Extract this month and last months data 
            list_current_month = list_multiple_embedded[0]
            list_previous_month = list_multiple_embedded[1]
            trending_check = [] #This check was added for the graphs created later. It's possible to have a spike without any driving product catalog numbers, this will ensure a graph is still created for those instances
            if len(trending_flagged) > 0: #Format event descriptions text file 
                file_ed = open(name_file_ed, 'a')
                file_ed.write('\n' + star + '\n' + star + '\n' + '\t\tTrending Device Code(s):' + '\n' + dash ) #Add header for trending device codes 
                file_ed.close() 
            for x in trending_flagged:
                file_ed = open(name_file_ed, 'a')
                file_ed.write('\n\t\t' + x) #Add trending device code to event descriptions text file
                trending_curr_data = list(filter(lambda z: z[trending_index] == x, list_current_month)) #Create a new list for complaints per trending device code using current month data
                trending_prev_data = list(filter(lambda z: z[trending_index] == x, list_previous_month)) #Create a new list for complaints per trending device code using previous month data
                products_curr = []
                products_prev = []
                for y in trending_curr_data:
                    products_curr.append(y[product_index]) #Make a list of all the product catalog numbers with complaints per this trending device code using current month data
                for y in trending_prev_data:
                    products_prev.append(y[product_index]) #Make a list of all the product catalog numbers with complaints per this trending device code using previous month data
                products_curr_set = [{a for a in products_curr}] #Create a set out of the list 
                self_check = 0 #Check if the trending device code does not have any driving product catalog numbers 
                for y in products_curr_set:
                    for z in y:
                        counts = 1 #index/count for the event description text file
                        #The amount of complaints will influence the threshold for determining if a product catalog number is a driver 
                        if products_curr.count(z) < 10: #Count how many complaints there are for this trending device code per product catalog number
                            if products_curr.count(z) - products_prev.count(z) > 1: #If the complaints are less than 10, an increase by 2 is a driver
                                file_ed.write('\n' + z) #Add the driver to the event descriptions text file
                                product_ed = list(filter(lambda o: o[product_index] == z, trending_curr_data)) #Create a new list of complaint details filtered for product catalog number for the this trending device code
                                for h in product_ed:
                                    if '\t' in h[ed_index]: #Format the event description text file
                                        h[ed_index] = h[ed_index].replace('\t', ' ')
                                    if 'It was reported that ' in h[ed_index]: #Format the event description text file
                                        h[ed_index] = h[ed_index].replace('It was reported that ','')
                                        capitalize_it = h[ed_index]
                                        capitalize = capitalize_it[0].upper() + capitalize_it[1:]
                                        h[ed_index] = capitalize
                                    file_ed.write('\n' + str(counts) + ') ' + h[ed_index]) #Add event description to the event description text file
                                    counts += 1
                                #Add details to the appropriate list
                                product_trending_tdc.append(x) #Trending Device Code
                                product_trending_p.append(z) #Product Catalog Number
                                product_trending_comp.append(products_curr.count(z)) #Product Catalog Number complaints
                                product_trending_diff.append(products_curr.count(z) - products_prev.count(z)) #Product Catalog Number complaint spike
                                file_ed.write('\n\n' + dash + '\n') #Formatting
                                self_check = 1 #Self check, this trending device code does have at least one product catolog number driver 
                        elif products_curr.count(z) < 20 and products_curr.count(z) >= 10: #Count how many complaints there are for this trending device code per product catalog number
                            if products_curr.count(z) - products_prev.count(z) > 2: #If the complaints are less than 20, an increase by 3 is a driver
                                file_ed.write('\n' + z) #Add the driver to the event descriptions text file
                                product_ed = list(filter(lambda o: o[product_index] == z, trending_curr_data)) #Create a new list of complaint details filtered for product catalog number for the this trending device code
                                for h in product_ed:
                                    if '\t' in h[ed_index]: #Format the event description text file
                                        h[ed_index] = h[ed_index].replace('\t', ' ')
                                    if 'It was reported that ' in h[ed_index]: #Format the event description text file
                                        h[ed_index] = h[ed_index].replace('It was reported that ','')
                                        capitalize_it = h[ed_index]
                                        capitalize = capitalize_it[0].upper() + capitalize_it[1:]
                                        h[ed_index] = capitalize
                                    file_ed.write('\n' + str(counts) + ') ' + h[ed_index]) #Add event description to the event description text file
                                    counts += 1 
                                #Add details to the appropriate list
                                product_trending_tdc.append(x) #Trending Device Code
                                product_trending_p.append(z) #Product Catalog Number
                                product_trending_comp.append(products_curr.count(z)) #Product Catalog Number complaints
                                product_trending_diff.append(products_curr.count(z) - products_prev.count(z)) #Product Catalog Number complaint spike
                                file_ed.write('\n\n' + dash + '\n') #Formatting
                                self_check = 1 #Self check, this trending device code does have at least one product catolog number driver 
                        elif products_curr.count(z) < 40 and products_curr.count(z) >= 20: #Count how many complaints there are for this trending device code per product catalog number
                            if products_curr.count(z) - products_prev.count(z) > 3: #If the complaints are less than 40, an increase by 4 is a driver
                                file_ed.write('\n' + z) #Add the driver to the event descriptions text file
                                product_ed = list(filter(lambda o: o[product_index] == z, trending_curr_data)) #Create a new list of complaint details filtered for product catalog number for the this trending device code
                                for h in product_ed:
                                    if '\t' in h[ed_index]: #Format the event description text file
                                        h[ed_index] = h[ed_index].replace('\t', ' ')
                                    if 'It was reported that ' in h[ed_index]: #Format the event description text file
                                        h[ed_index] = h[ed_index].replace('It was reported that ','')
                                        capitalize_it = h[ed_index]
                                        capitalize = capitalize_it[0].upper() + capitalize_it[1:]
                                        h[ed_index] = capitalize
                                    file_ed.write('\n' + str(counts) + ') ' + h[ed_index]) #Add event description to the event description text file
                                    counts += 1
                                #Add details to the appropriate list
                                product_trending_tdc.append(x) #Trending Device Code
                                product_trending_p.append(z) #Product Catalog Number
                                product_trending_comp.append(products_curr.count(z)) #Product Catalog Number complaints
                                product_trending_diff.append(products_curr.count(z) - products_prev.count(z)) #Product Catalog Number complaint spike
                                file_ed.write('\n\n' + dash + '\n') #Formatting
                                self_check = 1 #Self check, this trending device code does have at least one product catolog number driver 
                        elif products_curr.count(z) >= 40: #Count how many complaints there are for this trending device code per product catalog number
                            if ((products_curr.count(z) - products_prev.count(z))/products_curr.count(z))*100 >= 10: #If the complaints are 40 or greater, an increase by 10% is a driver
                                file_ed.write('\n' + z) #Add the driver to the event descriptions text file
                                product_ed = list(filter(lambda o: o[product_index] == z, trending_curr_data)) #Create a new list of complaint details filtered for product catalog number for the this trending device code
                                for h in product_ed:
                                    if '\t' in h[ed_index]: #Format the event description text file
                                        h[ed_index] = h[ed_index].replace('\t', ' ')
                                    if 'It was reported that ' in h[ed_index]: #Format the event description text file
                                        h[ed_index] = h[ed_index].replace('It was reported that ','')
                                        capitalize_it = h[ed_index]
                                        capitalize = capitalize_it[0].upper() + capitalize_it[1:]
                                        h[ed_index] = capitalize
                                    file_ed.write('\n' + str(counts) + ') ' + h[ed_index]) #Add event description to the event description text file
                                    counts += 1
                                #Add details to the appropriate list
                                product_trending_tdc.append(x) #Trending Device Code
                                product_trending_p.append(z) #Product Catalog Number
                                product_trending_comp.append(products_curr.count(z)) #Product Catalog Number complaints
                                product_trending_diff.append(products_curr.count(z) - products_prev.count(z)) #Product Catalog Number complaint spike
                                file_ed.write('\n\n' + dash + '\n') #Formatting
                                self_check = 1 #Self check, this trending device code does have at least one product catolog number driver 
                if self_check == 0: #If there is not a driving product catalog number, extract details for the future graph
                    trending_check.append(x) #Append the trending device code to a list specific to trending device codes without driving product catalog numbers
                    file_ed.write('\n\t' + 'Note: No Driving Product Catalog Numbers') #Format event descriptions text file
                    for y in products_curr_set: #If there are no driving product catalog numbers, add all the event descriptions for this trending device code to the text file
                        for z in y:
                            file_ed.write('\n' + z) #Format the event description text file
                            product_ed = list(filter(lambda o: o[product_index] == z, trending_curr_data)) #Create a new list of complaint details filtered for product catalog number for the this trending device code
                            for h in product_ed:
                                if '\t' in h[ed_index]: #Format the event description text file
                                    h[ed_index] = h[ed_index].replace('\t', ' ')
                                if 'It was reported that ' in h[ed_index]: #Format the event description text file
                                    h[ed_index] = h[ed_index].replace('It was reported that ','')
                                    capitalize_it = h[ed_index]
                                    capitalize = capitalize_it[0].upper() + capitalize_it[1:]
                                    h[ed_index] = capitalize
                                file_ed.write('\n' + str(counts) + ') ' + h[ed_index]) #Add event description to the event description text file
                                counts += 1
                file_ed.write('\n\n' + star + '\n') #Formatting
                file_ed.close()
            #Add the details from the new lists to the dictionary specific for trending device codes
            trending_products['Trending Device Code'] = product_trending_tdc
            trending_products['Product Catalog Number'] = product_trending_p
            trending_products['Complaints'] = product_trending_comp
            trending_products['Monthly Complaint Increase'] = product_trending_diff
            #Extract the infromation from the first dictionary to create a new dictionary. This dictionary will be more specific and only include product minors
            #The event descriptions text file will be updated simultaneously as the new dictionary informaiton is organized
            minor_tren = {} #New dictionary specific to product minors
            minor_flagged = spike['Product Minor Groups']
            #Create new lists for organizing data for the new dictionary
            minor_products_cat = []
            products_minor =[]
            minor_products_comp = []
            minor_products_diff = []
            minor_check = [] #This check was added for the graphs created later. It's possible to have a spike without any driving product catalog numbers, this will ensure a graph is still created for those instances
            if len(minor_flagged) > 0: #Format event descriptions text file 
                file_ed = open(name_file_ed, 'a')
                file_ed.write('\n' + star + '\n' + star + '\n' + '\t\tProduct Minor(s):' + '\n' + dash)
                file_ed.close()
            for x in minor_flagged:
                file_ed = open(name_file_ed, 'a')
                file_ed.write('\n\t\t' + x) #Add product minor to event descriptions text file
                minor_curr_data = list(filter(lambda z: z[minor_index] == x, list_current_month)) #Create a new list for complaints per product minor using current month data
                minor_prev_data = list(filter(lambda z: z[minor_index] == x, list_previous_month)) #Create a new list for complaints per product minor using previous month data
                minor_curr = []
                minor_prev = []
                for y in minor_curr_data:
                    minor_curr.append(y[product_index]) #Make a list of all the product catalog numbers with complaints per this product minor using current month data
                for y in minor_prev_data:
                    minor_prev.append(y[product_index]) #Make a list of all the product catalog numbers with complaints per this product minor using previous month data
                minor_curr_set = [{a for a in minor_curr}] #Create a set out of the list 
                self_check = 0 #Check if the product minor does not have any driving product catalog numbers
                for y in minor_curr_set:
                    for z in y:
                        counts = 1 #index/count for the event description text file
                        #The amount of complaints will influence the threshold for determining if a product catalog number is a driver 
                        if minor_curr.count(z) < 10:  #Count how many complaints there are for this product minor per product catalog number
                            if minor_curr.count(z) - minor_prev.count(z) > 1: #If the complaints are less than 10, an increase by 2 is a driver
                                file_ed.write('\n' + z) #Add the driver to the event descriptions text file
                                product_ed = list(filter(lambda o: o[product_index] == z, minor_curr_data)) #Create a new list of complaint details filtered for product catalog number for the this product minor
                                for h in product_ed:
                                    if '\t' in h[ed_index]: #Format the event description text file
                                        h[ed_index] = h[ed_index].replace('\t', ' ')
                                    if 'It was reported that ' in h[ed_index]: #Format the event description text file
                                        h[ed_index] = h[ed_index].replace('It was reported that ','')
                                        capitalize_it = h[ed_index]
                                        capitalize = capitalize_it[0].upper() + capitalize_it[1:]
                                        h[ed_index] = capitalize
                                    file_ed.write('\n' + str(counts) + ') ' + h[ed_index]) #Add event description to the event description text file
                                    counts += 1
                                #Add details to the appropriate list
                                minor_products_cat.append(x) #Product minor
                                products_minor.append(z) #Product Catalog Number
                                minor_products_comp.append(minor_curr.count(z)) #Product Catalog Number complaints
                                minor_products_diff.append(minor_curr.count(z) - minor_prev.count(z)) #Product Catalog Number complaint spike
                                file_ed.write('\n\n' + dash + '\n') #Formatting
                                self_check = 1 #Self check, this product minor does have at least one product catolog number driver 
                        elif minor_curr.count(z) < 20 and minor_curr.count(z) >= 10: #Count how many complaints there are for this product minor per product catalog number
                            if minor_curr.count(z) - minor_prev.count(z) > 2: #If the complaints are less than 20, an increase by 3 is a driver
                                file_ed.write('\n' + z) #Add the driver to the event descriptions text file
                                product_ed = list(filter(lambda o: o[product_index] == z, minor_curr_data)) #Create a new list of complaint details filtered for product catalog number for the this product minor
                                for h in product_ed:
                                    if '\t' in h[ed_index]: #Format the event description text file
                                        h[ed_index] = h[ed_index].replace('\t', ' ')
                                    if 'It was reported that ' in h[ed_index]: #Format the event description text file
                                        h[ed_index] = h[ed_index].replace('It was reported that ','')
                                        capitalize_it = h[ed_index]
                                        capitalize = capitalize_it[0].upper() + capitalize_it[1:]
                                        h[ed_index] = capitalize
                                    file_ed.write('\n' + str(counts) + ') ' + h[ed_index]) #Add event description to the event description text file
                                    counts += 1 
                                #Add details to the appropriate list
                                minor_products_cat.append(x) #Product minor
                                products_minor.append(z) #Product Catalog Number
                                minor_products_comp.append(minor_curr.count(z)) #Product Catalog Number complaints
                                minor_products_diff.append(minor_curr.count(z) - minor_prev.count(z)) #Product Catalog Number complaint spike
                                file_ed.write('\n\n' + dash + '\n') #Formatting
                                self_check = 1 #Self check, this product minor does have at least one product catolog number driver 
                        elif minor_curr.count(z) < 40 and minor_curr.count(z) >= 20: #Count how many complaints there are for this product minor per product catalog number
                            if minor_curr.count(z) - minor_prev.count(z) > 3: #If the complaints are less than 40, an increase by 4 is a driver
                                file_ed.write('\n' + z) #Add the driver to the event descriptions text file
                                product_ed = list(filter(lambda o: o[product_index] == z, minor_curr_data)) #Create a new list of complaint details filtered for product catalog number for the this product minor
                                for h in product_ed:
                                    if '\t' in h[ed_index]: #Format the event description text file
                                        h[ed_index] = h[ed_index].replace('\t', ' ')
                                    if 'It was reported that ' in h[ed_index]: #Format the event description text file
                                        h[ed_index] = h[ed_index].replace('It was reported that ','')
                                        capitalize_it = h[ed_index]
                                        capitalize = capitalize_it[0].upper() + capitalize_it[1:]
                                        h[ed_index] = capitalize
                                    file_ed.write('\n' + str(counts) + ') ' + h[ed_index]) #Add event description to the event description text file
                                    counts += 1
                                #Add details to the appropriate list
                                minor_products_cat.append(x) #Product minor
                                products_minor.append(z) #Product Catalog Number
                                minor_products_comp.append(minor_curr.count(z)) #Product Catalog Number complaints
                                minor_products_diff.append(minor_curr.count(z) - minor_prev.count(z)) #Product Catalog Number complaint spike
                                file_ed.write('\n\n' + dash + '\n') #Formatting
                                self_check = 1 #Self check, this product minor does have at least one product catolog number driver 
                        elif minor_curr.count(z) >= 40: #Count how many complaints there are for this product minor per product catalog number
                            if ((minor_curr.count(z) - minor_prev.count(z))/minor_curr.count(z))*100 >= 10:  #If the complaints are 40 or greater, an increase by 10% is a driver
                                file_ed.write('\n' + z) #Add the driver to the event descriptions text file
                                product_ed = list(filter(lambda o: o[product_index] == z, minor_curr_data)) #Create a new list of complaint details filtered for product catalog number for the this product minor
                                for h in product_ed:
                                    if '\t' in h[ed_index]: #Format the event description text file
                                        h[ed_index] = h[ed_index].replace('\t', ' ')
                                    if 'It was reported that ' in h[ed_index]: #Format the event description text file
                                        h[ed_index] = h[ed_index].replace('It was reported that ','')
                                        capitalize_it = h[ed_index]
                                        capitalize = capitalize_it[0].upper() + capitalize_it[1:]
                                        h[ed_index] = capitalize
                                    file_ed.write('\n' + str(counts) + ') ' + h[ed_index]) #Add event description to the event description text file
                                    counts += 1
                                #Add details to the appropriate list
                                minor_products_cat.append(x) #Product minor
                                products_minor.append(z) #Product Catalog Number 
                                minor_products_comp.append(minor_curr.count(z)) #Product Catalog Number complaints
                                minor_products_diff.append(minor_curr.count(z) - minor_prev.count(z)) #Product Catalog Number complaint spike
                                file_ed.write('\n\n' + dash + '\n') #Formatting
                                self_check = 1 #Self check, this product minor does have at least one product catolog number driver 
                if self_check == 0: #If there is not a driving product catalog number, extract details for the future graph
                    minor_check.append(x) #Append the product minor to a list specific to product minors without driving product catalog numbers
                    file_ed.write('\n\t' + 'Note: No Driving Product Catalog Numbers') #Format event descriptions text file
                    for y in minor_curr_set: #If there are no driving product catalog numbers, add all the event descriptions for this product minor to the text file
                        for z in y:
                            file_ed.write('\n' + z) #Format the event description text file
                            product_ed = list(filter(lambda o: o[product_index] == z, minor_curr_data)) #Create a new list of complaint details filtered for product catalog number for the this product minor
                            for h in product_ed:
                                if '\t' in h[ed_index]: #Format the event description text file
                                    h[ed_index] = h[ed_index].replace('\t', ' ')
                                if 'It was reported that ' in h[ed_index]: #Format the event description text file
                                    h[ed_index] = h[ed_index].replace('It was reported that ','')
                                    capitalize_it = h[ed_index]
                                    capitalize = capitalize_it[0].upper() + capitalize_it[1:]
                                    h[ed_index] = capitalize
                                file_ed.write('\n' + str(counts) + ') ' + h[ed_index]) #Add event description to the event description text file
                                counts += 1
                file_ed.write('\n\n' + star + '\n') #Formatting
                file_ed.close()
            #Add the details from the new lists to the dictionary specific for product minors
            minor_tren['Product Minor'] = minor_products_cat
            minor_tren['Product Catalog Number'] = products_minor
            minor_tren['Complaints'] = minor_products_comp
            minor_tren['Monthly Complaint Increase'] = minor_products_diff
            #Extract the infromation from the first dictionary to create a new dictionary. This dictionary will be more specific and only include product catalog numbers
            #The event descriptions text file will be updated simultaneously as the new dictionary informaiton is organized
            products_tren = {} #New dictionary specific to product catalog numbers
            product_flagged = spike['Product Catalog Numbers']
            #Create new lists for organizing data for the new dictionary
            minor_products_cat = []
            products_minor =[]
            minor_products_comp = []
            minor_products_diff = []
            if len(product_flagged) > 0: #Format event descriptions text file 
                file_ed = open(name_file_ed, 'a')
                file_ed.write('\n' + star + '\n' + star + '\n' + '\t\tProduct Catalog Number(s):' + '\n' + dash)
                file_ed.close()
            for x in product_flagged:
                file_ed = open(name_file_ed, 'a')
                file_ed.write('\n\t\t' + x) #Add product catalog number to event descriptions text file
                minor_curr_data = list(filter(lambda z: z[product_index] == x, list_current_month)) #Create a new list for complaints per product catalog number using current month data
                minor_prev_data = list(filter(lambda z: z[product_index] == x, list_previous_month)) #Create a new list for complaints per product catalog number using previous month data
                minor_curr = []
                minor_prev = []
                for y in minor_curr_data:
                    minor_curr.append(y[minor_index]) #Make a list using the product minor for the current months data. Every complaint for this product catalog number will add its minor to this list
                for y in minor_prev_data:
                    minor_prev.append(y[minor_index]) #Make a list using the product minor for the previous months data. Every complaint for this product catalog number will add its minor to this list
                minor_curr_set = [{a for a in minor_curr}] #Make a set of the product minors (there should only be one)
                for y in minor_curr_set:
                    for z in y:
                        counts = 1  #index/count for the event description text file
                        #The amount of complaints will influence the threshold for determining if a product catalog number had a complaint spike 
                        #The below if statement is left here incase a less than 10 complaint spike threshold is needed
                        if minor_curr.count(z) < 10: #Use the product minor list to count how many complaints the product catalog number had
                            if minor_curr.count(z) - minor_prev.count(z) > 2: #If the complaints are less than 10, an increase by 3 is a spike
                                for h in minor_curr_data:
                                    if '\t' in h[ed_index]: #Format the event description text file
                                        h[ed_index] = h[ed_index].replace('\t', ' ')
                                    if 'It was reported that ' in h[ed_index]: #Format the event description text file
                                        h[ed_index] = h[ed_index].replace('It was reported that ','')
                                        capitalize_it = h[ed_index]
                                        capitalize = capitalize_it[0].upper() + capitalize_it[1:]
                                        h[ed_index] = capitalize
                                    file_ed.write('\n' + str(counts) + ') ' + h[ed_index]) #Add event description to the event description text file
                                    counts += 1
                                #Add details to the appropriate list
                                minor_products_cat.append(x) #Product Catalog Number
                                products_minor.append(z) #Product minor
                                minor_products_comp.append(minor_curr.count(z)) #Product Catalog Number complaints (counted by using the minor list)
                                minor_products_diff.append(minor_curr.count(z) - minor_prev.count(z)) #Product Catalog Number complaint spike (counted by using the minor list)
                                file_ed.write('\n\n' + dash + '\n') #Formatting
                        elif minor_curr.count(z) < 20 and minor_curr.count(z) >= 10: #Use the product minor list to count how many complaints the product catalog number had
                            if minor_curr.count(z) - minor_prev.count(z) > 2: #If the complaints are less than 20, an increase by 3 is a spike
                                for h in minor_curr_data:
                                    if '\t' in h[ed_index]: #Format the event description text file
                                        h[ed_index] = h[ed_index].replace('\t', ' ')
                                    if 'It was reported that ' in h[ed_index]: #Format the event description text file
                                        h[ed_index] = h[ed_index].replace('It was reported that ','')
                                        capitalize_it = h[ed_index]
                                        capitalize = capitalize_it[0].upper() + capitalize_it[1:]
                                        h[ed_index] = capitalize
                                    file_ed.write('\n' + str(counts) + ') ' + h[ed_index]) #Add event description to the event description text file
                                    counts += 1
                                #Add details to the appropriate list
                                minor_products_cat.append(x) #Product Catalog Number
                                products_minor.append(z) #Product minor
                                minor_products_comp.append(minor_curr.count(z)) #Product Catalog Number complaints (counted by using the minor list)
                                minor_products_diff.append(minor_curr.count(z) - minor_prev.count(z)) #Product Catalog Number complaint spike (counted by using the minor list)
                                file_ed.write('\n\n' + dash + '\n') #Formatting
                        elif minor_curr.count(z) < 40 and minor_curr.count(z) >= 20: #Use the product minor list to count how many complaints the product catalog number had
                            if minor_curr.count(z) - minor_prev.count(z) > 3: #If the complaints are less than 40, an increase by 4 is a spike
                                for h in minor_curr_data:
                                    if '\t' in h[ed_index]: #Format the event description text file
                                        h[ed_index] = h[ed_index].replace('\t', ' ')
                                    if 'It was reported that ' in h[ed_index]: #Format the event description text file
                                        h[ed_index] = h[ed_index].replace('It was reported that ','')
                                        capitalize_it = h[ed_index]
                                        capitalize = capitalize_it[0].upper() + capitalize_it[1:]
                                        h[ed_index] = capitalize
                                    file_ed.write('\n' + str(counts) + ') ' + h[ed_index]) #Add event description to the event description text file
                                    counts += 1
                                #Add details to the appropriate list
                                minor_products_cat.append(x) #Product Catalog Number
                                products_minor.append(z) #Product minor
                                minor_products_comp.append(minor_curr.count(z)) #Product Catalog Number complaints (counted by using the minor list)
                                minor_products_diff.append(minor_curr.count(z) - minor_prev.count(z)) #Product Catalog Number complaint spike (counted by using the minor list)
                                file_ed.write('\n\n' + dash + '\n') #Formatting
                        elif minor_curr.count(z) >= 40: #Use the product minor list to count how many complaints the product catalog number had
                            if ((minor_curr.count(z) - minor_prev.count(z))/minor_curr.count(z))*100 >= 10: #If the complaints are 40 or greater, an increase by 10% is a spike
                                for h in minor_curr_data:
                                    if '\t' in h[ed_index]: #Format the event description text file
                                        h[ed_index] = h[ed_index].replace('\t', ' ')
                                    if 'It was reported that ' in h[ed_index]: #Format the event description text file
                                        h[ed_index] = h[ed_index].replace('It was reported that ','')
                                        capitalize_it = h[ed_index]
                                        capitalize = capitalize_it[0].upper() + capitalize_it[1:]
                                        h[ed_index] = capitalize
                                    file_ed.write('\n' + str(counts) + ') ' + h[ed_index]) #Add event description to the event description text file
                                    counts += 1
                                #Add details to the appropriate list
                                minor_products_cat.append(x) #Product Catalog Number
                                products_minor.append(z) #Product minor
                                minor_products_comp.append(minor_curr.count(z)) #Product Catalog Number complaints (counted by using the minor list)
                                minor_products_diff.append(minor_curr.count(z) - minor_prev.count(z)) #Product Catalog Number complaint spike (counted by using the minor list)
                                file_ed.write('\n\n' + dash + '\n') #Formatting
                file_ed.write('\n\n' + star + '\n') #Formatting
                file_ed.close()
            #Add the details from the new lists to the dictionary specific for Product Catalog Numbers
            products_tren['Product Catalog Number'] = minor_products_cat
            products_tren['Complaints'] = minor_products_comp
            products_tren['Monthly Complaint Increase'] = minor_products_diff

            #The next segments of code focus on creating the excel file and graphs for the flagged categories with complaint spikes
            cannots = ['\\','/',':','*','?','"','<','>','|'] #Symbols that cannot be in the name of a image file
            rrc = rr + '\\Complaint Summary.xlsx' #Name for excel summary file
            writer = pd.ExcelWriter(rrc, engine='xlsxwriter') #Create excel file
            q = [] 
            dq=pd.DataFrame(q) 
            dq.to_excel(writer,sheet_name='Graphs') #Add an empty dictionary to the excel file to create a new sheet
            writer.save() #Save the excel file
            wb = openpyxl.load_workbook(rrc) #Open the excel file using openpyxl for edditing 
            ws = wb.active #Active workbook / excel file
            str1 = 'Flagged Trending Device Codes with Driving Product Catalog Numbers' #Header for excel sheet for graphs
            str2 = 'Flagged Product Minors with Driving Product Catalog Numbers' #Header for excel sheet for graphs
            str3 = 'Flagged Product Catalog Numbers Independent of TDCs or PMs' #Header for excel sheet for graphs
            ws['A1'] = str1 #Add header to a cell
            ws['L1'] = str2 #Add header to a cell
            ws['W1'] = str3 #Add header to a cell
            wb.save(rrc) #Save the workbook
            #Trending device code graphs
            #Extract information from the trending device code dictionary 
            products_affected = trending_products['Product Catalog Number']
            graphs_trending = trending_products['Trending Device Code']
            graphs_trending_set = [{a for a in graphs_trending}]
            #Create a list of anchors (cells) for placing the graphs. Each excel column aligns to a the different header 
            anchorsa = ['A3','A29','A56','A83','A110','A137','A164','A191','A218','A245','A272','A299','A326','A353','A380','A407','A434','A461','A488','A515','A542','A569','A596','A623','A650','A677','A704','A731','A758','A785','A812']
            anchorsl = ['L3','L29','L56','L83','L110','L137','L164','L191','L218','L245','L272','L299','L326','L353','L380','L407','L434','L461','L488','L515','L542','L569','L596','L623','L650','L677','L704','L731','L758','L785','L812']
            anchorsw = ['W3','W29','W56','W83','W110','W137','W164','W191','W218','W245','W272','W299','W326','W353','W380','W407','W434','W461','W488','W515','W542','W569','W596','W623','W650','W677','W704','W731','W758','W785','W812']
            index_anchor = -1 #Index for the anchors list
            for x in graphs_trending_set:
                for y in x:
                    #New list for information needed for the graphs (axes, legend)
                    graphs_trending_p = [] #Legend 
                    graphs_trending_data = [] #Y-axis (one set of values eventually added to graphs_trending_p_data)
                    graphs_trending_p_data = [] #Y-axis
                    months = [] #X-axis 
                    index = -1
                    for z in list_multiple_embedded: #Index through the complaints per month
                        index += 1
                        list_trending = list(filter(lambda q: q[trending_index] == y,z)) #Create a new list for all complaints for this trending device code for this month
                        graphs_trending_data.append(len(list_trending)) #Track the number of complaints (y-axis)
                        months.append(month_names[str(all_months[index])]) #Track the month (x-axis)
                        index_2 = -1
                        graphs_trending_psub = []
                        for t in products_affected: #Index through the products flagged for complaint spike in the current the month 
                            index_2 += 1
                            if y == graphs_trending[index_2]: #Remember, products affected and graphs trending are the same length and every index aligns to the same complaint as the index of the other
                                list_products_aff = list(filter(lambda q: q[product_index] == t,list_trending)) #When the alignment is correct, make a list of all the complaints for a spiked product catalog number that is also for the trending device code
                                graphs_trending_psub.append(len(list_products_aff)) #Track the number of complaints (y-axis)
                        graphs_trending_p_data.append(graphs_trending_psub) #Consolidate y-axis values
                    index_2 = -1
                    for t in products_affected: #Add the product catalog numbers for the month to a list
                        index_2 += 1
                        if y == graphs_trending[index_2]:
                            graphs_trending_p.append(t) #Track the product catalog number (legend)
                    graphs_trending_p_data = list(map(list, zip(*graphs_trending_p_data))) #Change the configuration of the matrix (makes it easier to graph)
                    graphs_trending_p_data.append(graphs_trending_data) #Consolidate y-axis values
                    graphs_trending_p.append(y) #legend
                    index = -1
                    for z in graphs_trending_p_data:
                        index += 1
                        z = [t for t in z[::-1]]
                        graphs_trending_p_data[index] = z #Reverse the order of every list embedded in graphs_trending_p_data
                    months = [a for a in months[::-1]] #Reverse the order 
                    df = pd.DataFrame(graphs_trending_p_data, columns = months) #Create a DataFrame with the y-axis and x-axis data
                    plt.figure() #Create a figure
                    plt.plot(df.T) #Plot the transposed version of the dataframe
                    plt.legend([z for z in graphs_trending_p],loc='best') #Add the legend
                    plt.xlabel('Months') #Label the x-axis 
                    plt.ylabel('Complaints') #Label the y-axis
                    plt.title('%s' % y) #Title the graph by the trending device code
                    namefile = y 
                    for p in cannots: 
                        if p in namefile:
                            namefile = namefile.replace(p,' ') #Make sure the name of the trending device code does not have symbols not allowed in directories 
                    plt.savefig((rr + '\\' + namefile +'.png')) #Save the graph as an image file 
                    wb = openpyxl.load_workbook(rrc) #Open the excel file
                    ws = wb.active #Active worksheet
                    img = openpyxl.drawing.image.Image(rr + '\\' + namefile + '.png') #Open the image file of the graph
                    index_anchor += 1
                    img.anchor = anchorsa[index_anchor] #Add the anchor for the image 
                    ws.add_image(img) #Attach the image to the excel file
                    wb.save(rrc) #Save the excel file
                    #Resource: https://stackoverflow.com/questions/10888969/insert-image-in-openpyxl
                    #Extra resource for working with images and excel: https://stackoverflow.com/questions/33881280/how-do-i-insert-multiple-non-overlapping-images-using-openpyxl
            #Make graphs for trending device codes without driving product catalog numbers
            for x in trending_check:
                trending_check_data = []
                months = []
                index = -1
                for y in list_multiple_embedded:
                    index += 1
                    list_check_trending = list(filter(lambda z: z[trending_index] == x,y)) #Make a list of complaints for the trending device code from the month
                    trending_check_data.append(len(list_check_trending)) #Track the number of complaints (y-axis)
                    months.append(month_names[str(all_months[index])]) #Track the months (x-axis)
                trending_check_data = [a for a in trending_check_data[::-1]] #Reverse the order
                months = [a for a in months[::-1]] #Reverse the order
                plt.figure() #Create a figure
                plt.plot(months, trending_check_data) #Plot the data
                plt.xlabel('Months') #Label the x-axis
                plt.ylabel('Complaints') #Label the y-axis
                plt.title('%s (No Driving Product Catalog Numbers)' %x) #Title the graph by the trending device code
                namefile = x
                for p in cannots:
                    if p in namefile:
                        namefile = namefile.replace(p,' ') #Make sure the name of the trending device code does not have symbols not allowed in directories 
                plt.savefig((rr + '\\' + namefile + '.png')) #Save the graph as an image file 
                wb = openpyxl.load_workbook(rrc) #Open the excel file
                ws = wb.active #Active worksheet
                img = openpyxl.drawing.image.Image(rr + '\\' + namefile + '.png') #Open the image file of the graph
                index_anchor += 1
                img.anchor = anchorsa[index_anchor] #Add the anchor for the image 
                ws.add_image(img) #Attach the image to the excel file
                wb.save(rrc) #Save the excel file


            index_anchor = -1 #Index for the anchors list
            products_affected = minor_tren['Product Catalog Number']
            graphs_trending = minor_tren['Product Minor']
            graphs_trending_set = [{a for a in graphs_trending}]
            for x in graphs_trending_set:
                for y in x:
                    #New list for information needed for the graphs (axes, legend)
                    graphs_trending_p = [] #Legend 
                    graphs_trending_data = [] #Y-axis (one set of values eventually added to graphs_trending_p_data)
                    graphs_trending_p_data = [] #Y-axis
                    months = [] #X-axis 
                    index = -1
                    for z in list_multiple_embedded: #Index through the complaints per month
                        index += 1
                        list_trending = list(filter(lambda q: q[minor_index] == y,z)) #Create a new list for all complaints for this product minor for this month
                        graphs_trending_data.append(len(list_trending)) #Track the number of complaints (y-axis)
                        months.append(month_names[str(all_months[index])]) #Track the month (x-axis)
                        index_2 = -1
                        graphs_trending_psub = []
                        for t in products_affected: #Index through the products flagged for complaint spike in the current the month 
                            index_2 += 1
                            if y == graphs_trending[index_2]: #Remember, products affected and graphs trending are the same length and every index aligns to the same complaint as the index of the other
                                list_products_aff = list(filter(lambda q: q[product_index] == t,list_trending)) #When the alignment is correct, make a list of all the complaints for a spiked product catalog number that is also for the product minor
                                graphs_trending_psub.append(len(list_products_aff)) #Track the number of complaints (y-axis)
                        graphs_trending_p_data.append(graphs_trending_psub) #Consolidate y-axis values
                    index_2 = -1
                    for t in products_affected: #Add the product catalog numbers for the month to a list
                        index_2 += 1
                        if y == graphs_trending[index_2]:
                            graphs_trending_p.append(t) #Track the product catalog number (legend)
                    graphs_trending_p_data = list(map(list, zip(*graphs_trending_p_data))) #Change the configuration of the matrix (makes it easier to graph)
                    graphs_trending_p_data.append(graphs_trending_data) #Consolidate y-axis values
                    graphs_trending_p.append(y) #legend
                    index = -1
                    for z in graphs_trending_p_data:
                        index += 1 
                        z = [t for t in z[::-1]]
                        graphs_trending_p_data[index] = z #Reverse the order of every list embedded in graphs_trending_p_data
                    months = [a for a in months[::-1]] #Reverse the order 
                    pf = pd.DataFrame(graphs_trending_p_data, columns = months) #Create a DataFrame with the y-axis and x-axis data
                    plt.figure() #Create a figure
                    plt.plot(pf.T) #Plot the transposed version of the dataframe
                    plt.legend([z for z in graphs_trending_p],loc='best') #Add the legend
                    plt.xlabel('Months') #Label the x-axis 
                    plt.ylabel('Complaints') #Label the y-axis
                    plt.title('%s' % y) #Title the graph by the product minor
                    namefile = y
                    for p in cannots: 
                        if p in namefile: 
                            namefile = namefile.replace(p,' ') #Make sure the name of the product minor does not have symbols not allowed in directories 
                    plt.savefig((rr + '\\' + namefile +'.png')) #Save the graph as an image file 
                    wb = openpyxl.load_workbook(rrc) #Open the excel file
                    ws = wb.active #Active worksheet
                    img = openpyxl.drawing.image.Image(rr + '\\' + namefile + '.png') #Open the image file of the graph
                    index_anchor += 1
                    img.anchor = anchorsl[index_anchor] #Add the anchor for the image 
                    ws.add_image(img) #Attach the image to the excel file
                    wb.save(rrc) #Save the excel file

            #Make graphs for product minors without driving product catalog numbers
            for x in minor_check:
                minor_check_data = []
                months = []
                index = -1
                for y in list_multiple_embedded:
                    index += 1
                    list_check_minor = list(filter(lambda z: z[minor_index] == x,y)) #Make a list of complaints for the product minor from the month
                    minor_check_data.append(len(list_check_minor)) #Track the number of complaints (y-axis)
                    months.append(month_names[str(all_months[index])]) #Track the months (x-axis)
                minor_check_data = [a for a in minor_check_data[::-1]] #Reverse the order
                months = [a for a in months[::-1]] #Reverse the order
                plt.figure() #Create a figure
                plt.plot(months, minor_check_data) #Plot the data
                plt.xlabel('Months') #Label the x-axis
                plt.ylabel('Complaints') #Label the y-axis
                plt.title('%s (No Driving Product Catalog Numbers)' %x) #Title the graph by the product minor
                namefile = x
                for p in cannots:
                    if p in namefile:
                        namefile = namefile.replace(p,' ') #Make sure the name of the product minor does not have symbols not allowed in directories
                plt.savefig((rr + '\\' + namefile + '.png')) #Save the graph as an image file 
                wb = openpyxl.load_workbook(rrc) #Open the excel file
                ws = wb.active #Active worksheet
                img = openpyxl.drawing.image.Image(rr + '\\' + namefile + '.png') #Open the image file of the graph
                index_anchor += 1
                img.anchor = anchorsl[index_anchor] #Add the anchor for the image 
                ws.add_image(img) #Attach the image to the excel file
                wb.save(rrc) #Save the excel file

            index_anchor = -1 #Index for the anchors list
            products_affected = products_tren['Product Catalog Number']
            graphs_trending_set = [{a for a in products_affected}]
            for x in graphs_trending_set:
                for y in x:
                    #New list for information needed for the graphs (axes)
                    graphs_trending_data = [] #Y-axis 
                    months = [] #X-axis
                    index = -1
                    for z in list_multiple_embedded: #Index through the complaints per month
                        index += 1
                        list_trending = list(filter(lambda q: q[product_index] == y,z)) #Create a new list for all complaints for this product catalog number for this month
                        graphs_trending_data.append(len(list_trending)) #Track the number of complaints (y-axis)
                        months.append(month_names[str(all_months[index])]) #Track the month (x-axis)
                    graphs_trending_data = [a for a in graphs_trending_data[::-1]] #Reverse the order
                    months = [a for a in months[::-1]] #Reverse the order
                    pf = pd.DataFrame(graphs_trending_data, index = months) #Create a DataFrame with the y-axis and x-axis data
                    plt.figure() #Create a figure
                    plt.plot(pf) #Plot the dataframe
                    plt.xlabel('Months') #Label the x-axis 
                    plt.ylabel('Complaints') #Label the y-axis
                    plt.title('%s' % y) #Title the graph by the product catalog number
                    namefile = y
                    for p in cannots: 
                        if p in namefile: 
                            namefile = namefile.replace(p,' ') #Make sure the name of the product catalog number does not have symbols not allowed in directories 
                    plt.savefig((rr + '\\' + namefile +'.png')) #Save the graph as an image file 
                    wb = openpyxl.load_workbook(rrc) #Open the excel file
                    ws = wb.active #Active worksheet
                    img = openpyxl.drawing.image.Image(rr + '\\' + namefile + '.png') #Open the image file of the graph
                    index_anchor += 1
                    img.anchor = anchorsw[index_anchor] #Add the anchor for the image 
                    ws.add_image(img) #Attach the image to the excel file
                    wb.save(rrc) #Save the excel file

            #Use the functions defined above to create a text file for additional informaiton
            detail_hospital(data,list_current_month,month_names,month,star,dash,product_index,trending_index,minor_index,name_file_text) #Complaintant Facility 
            detail_lot(data,list_current_month,month_names,month,star,dash,product_index,trending_index,minor_index,name_file_text) #Lot
            detail_quantity(data,list_current_month,month_names,month,star,dash,product_index,trending_index,minor_index,name_file_text) #Quantity Affected
            trending_products = pd.DataFrame(trending_products) #Add the trending device code dictionary to a dataframe 
            trending_products.sort_values(by='Trending Device Code') #Group data by trending device code 
            minor_tren = pd.DataFrame(minor_tren) #Add the product minor dictionary to a dataframe
            minor_tren.sort_values(by='Product Minor') #Group data by product minor
            spike = pd.DataFrame(dict([(k,pd.Series(v)) for k,v in spike.items()])) #Add the spike dictionary to a dataframe and reorganize the data
            #Open excel file
            book = load_workbook(rrc)
            writer = pd.ExcelWriter(rrc, engine = 'openpyxl' )
            writer.book = book 
            #Create new excel sheets
            spike.to_excel(writer, sheet_name='Flagged Complaint Increases')
            trending_products.to_excel(writer, sheet_name='TDCs with Driving Products')
            minor_tren.to_excel(writer, sheet_name='PMs with Driving Products')
            #Save and close the excel file
            writer.save()
            writer.close()
            #Open excel file and format the tables imported from the dictionaries 
            wb = load_workbook(rrc)
            ws1 = wb['Flagged Complaint Increases']
            #Increase the column widths 
            ws1.column_dimensions['B'].width = 37
            ws1.column_dimensions['C'].width = 25
            ws1.column_dimensions['D'].width = 30
            ws1.column_dimensions['E'].width = 37
            ws1.column_dimensions['F'].width = 26
            ws1.column_dimensions['G'].width = 31
            ws1.column_dimensions['H'].width = 24
            ws1.column_dimensions['I'].width = 25
            ws1.column_dimensions['J'].width = 31
            #Add hyper links for the text files
            ws1['E20'].font = Font(color = 'FFFF0000', bold = True)
            ws1['E20'].hyperlink = name_file_text
            ws1['E20'].value = 'Additional Findings'
            ws1['F20'].font = Font(color = 'FF000080', bold = True)
            ws1['F20'].hyperlink = name_file_ed
            ws1['F20'].value = 'Event Descriptions'
            #Extra hyperlink option for a website link 
            #ws1['G20'].font = Font(color = 'FF008000', bold = True) ###*** 
            #ws1['G20'].hyperlink = 'https:///
            #ws1['G20'].value = 'BD Stat Tools Link'
            #Increase the column widths 
            ws2 = wb['TDCs with Driving Products']
            ws2.column_dimensions['B'].width = 37
            ws2.column_dimensions['C'].width = 24
            ws2.column_dimensions['D'].width = 12
            ws2.column_dimensions['E'].width = 27
            ws3 = wb['PMs with Driving Products']
            ws3.column_dimensions['B'].width = 37
            ws3.column_dimensions['C'].width = 24
            ws3.column_dimensions['D'].width = 12
            ws3.column_dimensions['E'].width = 27
            wb.save(rrc) #Save excel file 
            #Additional resources for formatting excel files 
            #https://stackoverflow.com/questions/8440284/setting-styles-in-openpyxl
            #https://stackoverflow.com/questions/8440284/setting-styles-in-openpyxl
            #https://stackoverflow.com/questions/8440284/setting-styles-in-openpyxl

            #Delete the graph image files
            for x in os.listdir(rr + '\\'):
                if x.endswith('.png'):
                    rri = rr + '\\' + x
                    os.remove(rri)
    
            print('\n\'Complaint Summary.xlsx\' has been added to %s.' % preference) #Print a confirmation that the excel summary file was created
    return

mainfunc() #Start the program

#Additional Resources 
##https://xlsxwriter.readthedocs.io/workbook.html
#https://datatofish.com/executable-pyinstaller/
#https://stackoverflow.com/questions/39077661/adding-hyperlinks-in-some-cells-openpyxl
