﻿# outsorsing
Скрипты для выгрузки отчетов в xlsx по аутсорсингу в группе компаний "МАТОРИН". 
share_folder_size_acl.py - подсчет размера общих папок на сервере, и списков доступа к ним (пользователи домена), сортировка по компаниям и пользователям, должно запускаться на сервере на котором распиоложены папки.

private_folder_size.py - подсчет размера личных папок на сервере (каждая папка соответствуеют имени учетной записи пользователя 
в домене) и подсчет размера. сортировка по компаниям и пользователям. 

lotus-notes_users.py - подключается к IBM LN серверу (с помощью COM соединения, то есть требуется настроенный клиент LN) и выгружает 
из указанной БД пользователей и сортирует их по компаниям и должностям.

domain_users.py - выгружает список пользователей из AD и сортирует их по компаниям и группам.
