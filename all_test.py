import pandas as pd

def pd_to_lists_sp(name_of_sheet):
	 
	df = pd.read_excel('spare_part_list.xlsx', sheet_name=name_of_sheet)# can also index sheet by name or fetch all sheets
	''' df[['Item No.','Weight/kg','EXW Price/EURO','Description']].values.tolist() convetrs columns of df to lists of list'''
	return df['Match Code'].tolist(), df[['Item No.','Weight/kg','EXW Price/EURO','Description']].values.tolist(),\
	df['Item No.'].tolist(),df[['Match Code','Weight/kg','EXW Price/EURO','Description']].values.tolist()

def pd_to_list_pn_sds(name_of_sheet):
	df = pd.read_excel('spare_part_list.xlsx', sheet_name=name_of_sheet)
	return df['Project'].tolist(), df['Location'].tolist()

def pd_to_list_pn_other(name_of_sheet):
	df = pd.read_excel('spare_part_list.xlsx', sheet_name=name_of_sheet)
	return df['Customer'].tolist(), df[['Project_no','City']].values.tolist() #  df[['Project',City']].values.tolist(): Keyerror 

def pd_to_lists_agent(name_of_sheet):
	df = pd.read_excel('spare_part_list.xlsx', sheet_name=name_of_sheet)
	return df['Short Name'].tolist(),df[['Company','Person','Address','Telefone','E-Mail']].values.tolist()

def add_signature(ll, signature):
	for l in ll:
		l.append(signature)

def add_4_lists(list1, list2, list3, list4):
	return list1 + list2  + list3 + list4

def add_5_lists(list1, list2, list3, list4, list5):
	return list1 + list2  + list3 + list4 + list5 

def lists_to_dict(list1, list2):
	return dict(zip(list1, list2))

def lists_to_dict_m(l, ll):
	result = dict()
	'''multiple values to a key using list'''
	for elem, elem_l in zip(l, ll):
		'''already have a value to the key'''
		if elem in result:
			result[elem].append(elem_l)
		'''no value hashed to the key yet'''
		else:
			result[elem] = [elem_l]
	return result

def split_to_lists(list_input):
	list1= list()
	list2 = list()
	for element in list_input:
		list_temp = element.split()
		list1.append(list_temp[0])
		list2.append(list_temp[1])
	return list1, list2

def combine_to_ll(list_pn, list_loc):
	result = list()
	for pn, loc in zip(list_pn, list_loc):
		temp = list()
		temp.append(pn)
		temp.append(loc)
		result.append(temp)
	return result

def search_sp():
	while True:
		str_input = input("Enter the match code or item number or type q to return to the upper level memu, please:\n")
		print("\n")
		if str_input == "q":
			break
		elif str_input in dict_match:
			print(dict_match[str_input])
		elif str_input in dict_item:
			print(dict_item[str_input])
		else:
			print("It's not an Amepa SP!")
		print("\n")

def search_pn():
	while True:
		str_input = input("Enter the project name or type q to return to the upper level memu, please:\n")
		print("\n")
		if str_input == "q":
			break
		elif str_input in dict_name:
			for elem in dict_name[str_input]:
				print(elem)
				print("\n")
		else:
			print("It's not a Amepa Project!")
		print("\n")	

def search_agent():
	while True:
		str_input = input("Enter the agent company short name or type q to return to the upper level memu, please:\n")	
		if str_input == "q":
			break
		elif str_input in dict_agent:
			print(dict_agent[str_input])
		else:
			print("It's not a Amepa Project!")
		print("\n")	


def search():
	while True:
		str_input_num = input("Enter 1 for SP, enter 2 for project no., enter 3 for agent info or type q to quit\n")
		print("\n")
		if str_input_num == "1":
			search_sp()
		elif str_input_num == "2":
			search_pn()
		elif str_input_num == "3":
			search_agent()
		elif str_input_num == "q":
			break
		else:
			print("What you typed is not a recognizable option in this memu, please retype!")


if __name__ == '__main__':
	
	match_l_esd,ll_match_esd, item_l_esd, ll_item_esd = pd_to_lists_sp("ESD-SP")
	match_l_tsd1,ll_match_tsd1, item_l_tsd1, ll_item_tsd1 = pd_to_lists_sp("TSD 1.0-SP")
	match_l_tsd2,ll_match_tsd2, item_l_tsd2, ll_item_tsd2 = pd_to_lists_sp("TSD 2.0-SP")
	match_l_srm,ll_match_srm, item_l_srm, ll_item_srm = pd_to_lists_sp("SRM-SP")
	match_l_ofm,ll_match_ofm, item_l_ofm, ll_item_ofm = pd_to_lists_sp("OFM-SP")
	

	add_signature(ll_match_esd, "ESD")
	add_signature(ll_item_esd, "ESD")
	add_signature(ll_match_tsd1, "TSD1.0")
	add_signature(ll_item_tsd1, "TSD1.0")
	add_signature(ll_match_tsd2, "TSD2.0")
	add_signature(ll_item_tsd2, "TSD2.0")
	add_signature(ll_match_srm, "SRM")
	add_signature(ll_item_srm, "SRM")
	add_signature(ll_match_ofm, "OFM")
	add_signature(ll_item_ofm, "OFM")

	all_l_match = add_5_lists(match_l_esd, match_l_tsd1, match_l_tsd2, match_l_srm, match_l_ofm)
	all_ll_match = add_5_lists(ll_match_esd, ll_match_tsd1, ll_match_tsd2, ll_match_srm, ll_match_ofm)
	all_l_item = add_5_lists(item_l_esd, item_l_tsd1, item_l_tsd2, item_l_srm, item_l_ofm)
	all_ll_item = add_5_lists(ll_item_esd, ll_item_tsd1, ll_item_tsd2, ll_item_srm, ll_item_ofm)


	dict_match = lists_to_dict(all_l_match, all_ll_match)
	dict_item = lists_to_dict(all_l_item, all_ll_item)

	
	raw_project_list_tsd, location_list_tsd = pd_to_list_pn_sds("ESD-Projects")
	raw_project_list_esd, location_list_esd = pd_to_list_pn_sds("TSD-Projects")
	project_name_list_srm, ll_srm = pd_to_list_pn_other("SRM-Projects")
	project_name_list_ofm, ll_ofm = pd_to_list_pn_other("OFM-Projects")

	project_no_list_tsd, project_name_list_tsd = split_to_lists(raw_project_list_tsd)
	project_no_list_esd, project_name_list_esd = split_to_lists(raw_project_list_esd)

	ll_tsd = combine_to_ll(project_no_list_tsd, location_list_tsd)
	ll_esd = combine_to_ll(project_no_list_esd, location_list_esd)

	all_name_l = add_4_lists(project_name_list_tsd, project_name_list_esd, project_name_list_srm, project_name_list_ofm)
	all_name_ll = add_4_lists(ll_tsd, ll_esd, ll_srm, ll_ofm)

	dict_name = lists_to_dict_m(all_name_l, all_name_ll)

	agent_l, info_ll = pd_to_lists_agent("Agents")
	dict_agent = lists_to_dict(agent_l, info_ll)

	search()