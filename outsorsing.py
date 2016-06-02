# -*- coding: utf-8 -*-

import ldap
import sys
import locale
import codecs
import os
import win32security
import ntsecuritycon


def init_ldap():
	LDAP_URL = "ldap://172.20.20.5"
	USERNAME = "obruchev@matorin.local"
	PASSWORD = "OuKs1911"
	
	ad = ldap.initialize(LDAP_URL)
	ad.simple_bind_s(USERNAME, PASSWORD)
	ad.protocol_version = ldap.VERSION3
	ad.set_option(ldap.OPT_REFERRALS,0)

	return ad

def query_ldap (ad, group):
	basedn = "OU=MATORIN,DC=matorin,DC=local"
	scope = ldap.SCOPE_SUBTREE
	print(type(group), group)
	#если группа у пользоватлея основная, то нужно делать другой запрос
	if group == "Пользователи домена":
		filterexp = "(&(objectCategory=person)(objectClass=user)(primaryGroupID=513))"
	else :
		filterexp = "(&(objectCategory=person)(objectClass=user)(memberOf=CN="+group+",CN=Users,DC=matorin,DC=local))"
	#поиск по группе активных пользователей
	#filterexp = "(&(objectCategory=person)(objectClass=user)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(memberOf=CN=Only_internet,CN=Users,DC=matorin,DC=local))"
	#поиск всех активных пользователей
	#filterexp = "(&(objectCategory=person)(objectClass=user)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))"
	
	attrlist = None
	results = ad.search_s(basedn, scope, filterexp, attrlist)

	fullname = []
	company_n = []
	accountname = []

	print (filterexp.decode("utf-8"))
	for result in results:
		fullname.append(result[1]["cn"][0].decode("utf-8"))
		accountname.append(result[1]["sAMAccountName"][0].decode("utf-8"))
		try:
			company_n.append(result[1]["company"][0].decode("utf-8"))
		except KeyError:
			company_n.append(0)
    
	return fullname, company_n, accountname




def get_dir_size(path):
	DIR_EXCLUDES = set(['.', '..'])
	MASK = win32con.FILE_ATTRIBUTE_DIRECTORY | win32con.FILE_ATTRIBUTE_SYSTEM
	REQUIRED = win32con.FILE_ATTRIBUTE_DIRECTORY
	FindFilesW = win32file.FindFilesW
	total_size = 0
	try:
		items = FindFilesW(path + r'\*')
	except pywintypes.error, ex:
		return total_size
	for item in items:
		total_size += item[5]
		if (item[0] & MASK == REQUIRED):
			name = item[8]
			if name not in DIR_EXCLUDES:
				total_size += get_dir_size(path + '\\' + name)
	return total_size

def main(argv):
	#fullname = []
	##company = []
	#accountname = []
	#решаем проблему с кодировкой
	report_path = "C:\\scripts\\outsorsing"
	sys.stdout = codecs.getwriter('cp866')(sys.stdout,'replace')
	reload(sys)
	sys.setdefaultencoding('utf-8')

	group_domain = "Пользователи домена"
	groups_farm = ["Пользователи RDSFARM", "Пользователи квартплаты", "Only_internet"]

	print sys.getdefaultencoding()
	print locale.getpreferredencoding()
	print sys.stdout.encoding
	ad = init_ldap()
	
	fullname, company, accountname = query_ldap (ad,group_domain)

	uniq_company = list(set(company))

	for comp in uniq_company:
		count = 0
		filename = group_domain + ' ' + str(comp) + ".txt"
		filename = filename.decode("utf-8")
		fullpath = report_path + filename
		if not os.path.exists(filename):
			f = open(fullpath,'wt')
		else: 
			f = open(fullpath,'at')
		for i in xrange(len(fullname)):
			if company[i] == comp:
				count+=1
				f.write(accountname[i].rjust(20) + '\t' + fullname[i]  + '\n')
		f.write("\n===========" + '\n' + "Количество = " + str(count) + '\n')
		f.close()

	for group in groups_farm:
		fullname, company, accountname = query_ldap (ad,group)
		for comp in uniq_company:
			count = 0
			filename = groups_farm[0] + ' ' + str(comp) + ".txt"
			filename = filename.decode("utf-8")
			fullpath = report_path + filename
			if not os.path.exists(fullpath):
				f = open(fullpath,'wt')
			else: 
				f = open(fullpath,'at')
			for i in xrange(len(fullname)):
				if company[i] == comp :
					count+=1
					f.write(accountname[i].rjust(20) + '\t' + fullname[i]  + '\n')
			f.write("\n===========" + '\n' + "Количество = " + str(count) + '\n')
			f.close()

	#print(get_dir_size("c:\\Папка новая\\"))

	#for i in range(len(fullname)):
	#	print (fullname[i])
	#	print (company[i])
	#	print (accountname[i])

	#for x in uniq_company:
	#	print x
	
	ad.unbind_s()
	return 0
if __name__ == '__main__':
    sys.exit(main(sys.argv))	
