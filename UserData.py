data = [       
    {"ordinal_number": 1, "username": "olive", "name1": "Oliver Bregneberg", "Mail": "ob@timevat.com"}, 
   {"ordinal_number": 2, "username": "Timevat", "name1": "Sofus Lynge", "Mail": "sl@timevat.com"},    
   {"ordinal_number": 3, "username": "Xuan", "name1": "Xuan Vu", "Mail": "xv@timevat.com"}
   
   ]

def find_username_and_name(ordinal_number):
    for entry in data:
        if entry["ordinal_number"] == ordinal_number:
            username = entry["username"]
            name1 = entry["name1"]
            Mail = entry["Mail"]
            
            return username, name1, Mail

ordinal_input = int(input("Enter the ordinal number:\n \n 1. = Oliver\n 2. = Sofus\n 3. = Xuan\n \nAnswer = "))
username, name1, Mail = find_username_and_name(ordinal_input)
print("Username:", username)
print("Name:", name1)