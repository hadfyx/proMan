#############################################################################
'''
proMan by Andrej Perfilov [perfilov3d.com]

This project was rushed so please bear with me here... :)

Things yet to fix:
- split up into modules for easier maintenance
- lots of refactoring...
- implement a query class to handle all database communications
  (currently all over the place and using both sqlite and QtSql query methods)
- remote control (cmd.txt) update trigger mostly not working
- create a proper password routine
- don't use globals

Description:

    proMan (Project Manager) is a multi-user database application.
    All project data is kept in a single database and is accessible
    from any machine on the local network. Facilitates simultaneous
    updates and queries from multiple users.
    Written in Python 3.5 and Qt 5.6.

Features:

    - User view (user's ongoing projects)
    - Global view (all ongoing projects in the office)
    - Status filter (In Progress, Completed, etc.)
    - Email integration (metadata)
    - Record search (by project name, user, type, etc.)
    - Quick navigation (Right Click -> Go to Project)
    - Project history (Date, Status change, User)
    - Daily auto backup
    
'''
#############################################################################


import sys, os, sqlite3, subprocess, time, pathlib, shutil, glob, ast
from PyQt5 import QtCore, QtSql, QtGui
from PyQt5.QtWidgets import QApplication, QMainWindow, QDialog, QHeaderView, QFileDialog, \
                            QMessageBox, QMenu, qApp, QTreeWidgetItem
from PyQt5.uic import loadUi
from email.header import decode_header

# GLOBALS #
currUser = ''
dbPath = ''
mailList = []
root = os.getenv('APPDATA')
setDir = root + '\\proMan'
dbIni = setDir + '\\db_loc.ini'
userIni = setDir + '\\user.ini'
appIni = setDir + '\\app.ini'
mailIni = setDir + '\\mail.ini'
updateBat = setDir + '\\update.bat'
appName = 'proMan v1.3.1' + ' - '
# GLOBALS #


class recInfo():
    Id = ''
    props = ''
    user = ''
    priority = 50
    project = ''
    project_path = ''
    nr = ''
    Type = ''
    cad = ''
    _3d2d = ''
    rend = ''
    rend_path = ''
    post = ''
    rev = 0
    created = ''
    notes = ''
    status = 'In Progress'
    home = ''
    
  
class mainWindow(QMainWindow):
    '''
    ### COLUMN LIST ###
    0 - ID
    1 - Props
    2 - User
    3 - Priority
    4 - Project
    5 - Project_path
    6 - Nr
    7 - Type
    8 - CAD
    9 - DD # 3D-2D
    10 - Rend
    11 - Rend_path
    12 - Post
    13 - Rev
    14 - Created
    15 - Notes
    16 - Status
    17 - Home
    18 - Info_A
    19 - Info_B
    20 - Info_C
    21 - Info_D
    22 - Info_E
    23 - Info_F
    24 - Info_G
    25 - Info_H
    '''
    def __init__(self):
        QMainWindow.__init__(self)
        self.ui = loadUi('proMan_main.ui')
        self.ui.show()

        self.currColumn = False
        self.currRow = ''
        self.userSetFlag = False
        self.newStatus = recInfo.status
        self.historyList = []
        
        action = self.ui.menuFile.addAction('&Exit')
        action.triggered.connect(qApp.quit)
        self.ui.closeEvent = self.closeEvent # minimize instead of closing
        
        self.ui.but_settings.clicked.connect(self.openSettings)
        self.ui.but_new.clicked.connect(self.openRecord)
        self.ui.but_del.clicked.connect(self.delRecord)
        self.ui.lin_search.textChanged.connect(self.searchRecord)
        #self.ui.tbl_View.horizontalHeader().sortIndicatorChanged.connect(self.getColumn)
        self.ui.tbl_View.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.ui.tbl_View.customContextMenuRequested.connect(self.popUp)
        self.ui.chb_Global.stateChanged.connect(self.viewFilter)
        self.ui.cbx_user.currentIndexChanged.connect(self.filterRows)
        self.ui.cbx_user.currentIndexChanged.connect(self.displayMail)
        self.ui.cbx_status.currentIndexChanged.connect(self.filterRows)
        self.ui.chb_Mail.stateChanged.connect(self.showMail)
        self.ui.but_setUser.clicked.connect(self.setUser)
            
        if self.getSettings():
            self.setupDbView()
            self.setFirstUser()
            self.filterRows()
            self.recordMail()
            self.showMail()
            self.displayMail()
            if currUser:
                recInfo.user = currUser
                print(recInfo.user)
        
            self.mailWatch()
    
    def mailWatch(self):
        if mailList:
            self.mailWatcher = QtCore.QFileSystemWatcher() # watch for file changes
            self.mailWatcher.addPath(mailList[0])
            self.mailWatcher.directoryChanged.connect(self.recordMail)
    
    def showMail(self):
        if self.ui.chb_Mail.isChecked():
            self.ui.tw_Mail.show()
        else:
            self.ui.tw_Mail.hide()
    
    def lockCheck(self, file):
        test = False
        max_try = 5
        for i in range(max_try):
            try:
                with open(file, 'rb') as _:
                    test = True
                    break
            except:
                time.sleep(1)
        return test
    
    def getMail(self, msg_Inbox):
        allFiles = []
        msg_list = []
        
        def decodeMail(msg): # decode MIME data if present
            s = msg.split()
            msg_lst = []
            for i in s:
                if i:
                    new_str = i
                    bytes, encoding = decode_header(i)[0]
                    if encoding:
                        new_str = bytes.decode(encoding)
                    msg_lst.append(new_str)
            return ' '.join(msg_lst)
        
        def msgInfo(msg):
            get_list = [r'From: ',r'Subject: ',r'Date: ']
            info_list = []
            
            if os.path.exists(msg):
                if self.lockCheck(msg): # check if file is being written
                    with open(msg, 'rt', encoding='latin-1') as file:
                        for word in get_list:
                            result = ''
                            for line in file:
                                if line.startswith(word):
                                    result = decodeMail(line) # decode MIME data if present
                                    if word == r'Subject: ':
                                        info_list.append(result.rstrip().replace(r'Subject: ',''))
                                    else:
                                        info_list.append(result.rstrip())
                                    file.seek(0) # go to first line, in case the order of get_list items in the file is different
                                    break
                                else: continue
                                break
                            if not result: # if nothing found
                                info_list.append(word + 'none')
            return info_list
        
        working_dir = os.getcwd() # catch current working directory
        for search_dir in msg_Inbox:
            if search_dir:
                os.chdir(search_dir) # change current working directory
                files = filter(os.path.isfile, os.listdir(search_dir))
                files = [os.path.join(search_dir, f) for f in files] # add path to each file
                allFiles += files
        allFiles.sort(key=lambda x: os.path.getctime(x))
        os.chdir(working_dir) # restore current working directory
        
        count = 0
        for msg in reversed(allFiles):
            count += 1
            m = ''
            if count > 10:
                break
            else:
                m = msgInfo(msg)
                if m: msg_list.append(m)
        return [len(allFiles), msg_list]
    
    def recordMail(self):
        try:
            if not mailList:
                mail = [' Not Connected! ',[]]
            else:
                mail = self.getMail(mailList) # getMail sometimes fails when removing messages
            query = QtSql.QSqlQuery()
            query.prepare("UPDATE Users SET Mail=:ml WHERE User=:us")
            query.bindValue(':ml', str(mail))
            query.bindValue(':us', currUser)
            query.exec_()
        except: print('warning: recordMail failed!')
    
    def displayMail(self):
        if recInfo.user: # check if there's a row selection. recInfo.user gets set on row selection, if None - no selection
            if not self.ui.chb_Global.isChecked():
                user = self.ui.cbx_user.currentText()
            else:
                user = recInfo.user
            con = sqlite3.connect(dbPath)
            with con:
                cur = con.cursor()
                cur.execute('SELECT * FROM Users WHERE User=?', (user,))
                mailData = cur.fetchall()
            mailNum = '0'
            mail = []
            try: 
                m = ast.literal_eval(mailData[0][1]) # convert mail string to list
                mailNum = str(m[0])
                mail = m[1]
            except: mailNum = ' Not Connected! '
            self.ui.lab_userMail.setText(user + ' - Inbox(' + mailNum + ')')
            self.ui.tw_Mail.clear()
            if mail:
                for i in mail:
                    topItem = QTreeWidgetItem()
                    topItem.setText(0, i[1]) # subject as top level item
                    for j in (0,2): # 0 - from, 2 - date
                        item = QTreeWidgetItem()
                        item.setText(0, i[j])
                        topItem.addChild(item)
                    self.ui.tw_Mail.addTopLevelItem(topItem)
    
    def setUser(self):
        if self.userSetFlag:
            global currUser
            currUser = self.ui.cbx_user.currentText()
            if currUser:
                u = open(userIni, 'w')
                u.write(currUser)
                u.close()
            self.filterRows()
            self.ui.setWindowTitle(appName + '[ ' + currUser + ' ]')
        
    def remoteControl(self):
        
        def most_recent_file(path):
            glob_pattern = os.path.join(path, '*.exe')
            return max(glob.iglob(glob_pattern), key=os.path.getctime)
        
        control = os.path.dirname(dbPath) + '/cmd.txt'
        if os.path.exists(control):
            d = open(control, 'r')
            command = d.readline()
            d.close()
            if command == 'close':
                self.qApp.quit
            elif command == 'update':
                folder = str(pathlib.Path(dbPath).parents[1]) + r'update'
                file = most_recent_file(folder)
                if os.path.exists(updateBat):
                    b = open(updateBat, 'w')
                    b.write('Start' + ' "" ' + '"' + file + '"')
                    b.close()
                os.startfile(updateBat)
                time.sleep(3)
                self.qApp.quit
                
    
    def backUp(self):
        folder = str(pathlib.Path(dbPath).parents[1]) + 'backup'
        if not os.path.exists(folder):
            os.makedirs(folder)
        file = os.path.basename(dbPath)
        dest = folder + '\\' + time.strftime('%Y-%m-%d') + '_' + file
        if not os.path.exists(dest):
            print ('backup created',dest)
            shutil.copyfile(dbPath, dest)
    
    def setFirstUser(self):
        index = self.ui.cbx_user.findText(currUser, QtCore.Qt.MatchFixedString) # set the user
        if index >= 0:
            self.ui.cbx_user.setCurrentIndex(index)
        self.userSetFlag = True
            
    def getSettings(self):
        
        def createFiles(): # make sure files exist
            d = open(dbIni, 'a')
            d.close()
            u = open(userIni, 'a')
            u.close()
            a = open(appIni, 'a')
            a.close()
            m = open(mailIni, 'a')
            m.close()
            u = open(updateBat, 'a')
            u.close()
        
        if not os.path.exists(setDir):
            os.makedirs(setDir)
            createFiles()
            self.openSettings()
            self.ui.setWindowTitle(appName + 'No Database Loaded!')
            return False
        else:
            global dbPath
            global currUser
            global mailList
            createFiles()
            d = open(dbIni, 'r')
            dbPath = d.readline()
            d.close()
            u = open(userIni, 'r')
            currUser = u.readline()
            u.close()
            m = open(mailIni, 'r')
            mailList = m.readline().split(',')
            mailList = list(filter(None, mailList))
            m.close()
            if os.path.exists(dbPath):
                return True
            else:
                self.openSettings()
                self.ui.setWindowTitle(appName + 'No Database Loaded!')
                return False

    def addHistory(self):
        try:
            con = sqlite3.connect(dbPath)
            with con:
                cur = con.cursor()
                cur.execute('SELECT * FROM Projects WHERE ID=?', (recInfo.Id,))
                data = cur.fetchall()
                props = data[0][1]
                if props == None:
                    props = ''
                user = data[0][2]
                status = data[0][16]
            if status != 'In Progress': # make project live
                if user != recInfo.user:
                    reply = QMessageBox.question(self, 'Change Status?','Set this project In Progress?', QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                    if reply == QMessageBox.Yes:
                        query = QtSql.QSqlQuery()
                        query.prepare("UPDATE Projects SET Status=:st WHERE ID=:id")
                        self.newStatus = 'In Progress'
                        query.bindValue(':st', self.newStatus)
                        query.bindValue(':id', recInfo.Id)
                        query.exec_()
            if user != recInfo.user or status != self.newStatus:
                oldHist = props
                newHist = oldHist + time.strftime('%Y-%m-%d') + ' - ' + self.newStatus + ' - ' + recInfo.user + ','
                query = QtSql.QSqlQuery()
                query.prepare("UPDATE Projects SET Props=:pr WHERE ID=:id")
                query.bindValue(':pr', newHist)
                query.bindValue(':id', recInfo.Id)
                query.exec_()
                #self.sourceModel.select()
        except: print ('Could not update history')

    def popUp(self, pos):
        
        def changeStatus(stat):
            self.newStatus = stat
            self.addHistory()
            query = QtSql.QSqlQuery()
            query.prepare("UPDATE Projects SET Status=:st, Priority=:pr WHERE ID=:id")
            query.bindValue(':st', stat)
            query.bindValue(':pr', 50) # reset pr-ty
            query.bindValue(':id', recInfo.Id)
            query.exec_()
            self.updateChanges()
        
        table = self.ui.tbl_View
        if table.selectionModel().selection():
            #recInfo.Id = table.model().index(self.currRow, 0).data()
            menu = QMenu()
            clickProj = menu.addAction('Go to Project')
            clickRend = menu.addAction('Go to Render')
            menu.addSeparator()
            statusMenu = menu.addMenu('Mark as...')
            inPr = statusMenu.addAction('In Progress')
            comp = statusMenu.addAction('Completed')
            statusMenu.addSeparator()
            other = statusMenu.addMenu('Other...')
            sign = other.addAction('Signed Off')
            menu.addSeparator()
            clickProp = menu.addAction('Properties')
            action = menu.exec_(table.mapToGlobal(pos))
            if action == inPr:
                changeStatus('In Progress')
            elif action == comp:
                changeStatus('Completed')
            elif action == sign:
                reply = QMessageBox.question(self, 'Sign Off',"Sign Off selected project?", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                if reply == QMessageBox.Yes:
                    changeStatus('Signed Off')
            elif action == clickProj:
                self.goPlaces(5)
            elif action == clickRend:
                self.goPlaces(11)
            elif action == clickProp:
                
                # fetch parameters
                con = sqlite3.connect(dbPath)
                with con:
                    cur = con.cursor()
                    cur.execute('SELECT * FROM Projects WHERE ID=?', (recInfo.Id,))
                    data = cur.fetchall()
                    recInfo.props = data[0][1]
                    if recInfo.props == None: #check if no history
                        self.historyList = []
                    else:
                        self.historyList = (recInfo.props).split(',')
                    recInfo.user = data[0][2]
                    recInfo.project = data[0][4]
                    recInfo.project_path = data[0][5]
                    recInfo.nr = data[0][6]
                    recInfo.Type = data[0][7]
                    recInfo.rend_path = data[0][11]
                    recInfo.status = data[0][16]
                    self.newStatus = recInfo.status
                    recInfo.home = data[0][17]
                '''
                recInfo.props = (table.model().index(self.currRow, 1).data()).split(',')
                recInfo.user = table.model().index(self.currRow, 2).data()
                recInfo.project = table.model().index(self.currRow, 4).data()
                recInfo.project_path = table.model().index(self.currRow, 5).data()
                recInfo.nr = table.model().index(self.currRow, 6).data()
                recInfo.Type = table.model().index(self.currRow, 7).data()
                recInfo.rend_path = table.model().index(self.currRow, 11).data()
                recInfo.home = table.model().index(self.currRow, 17).data()
                '''
                # fetch parameters
                
                self.newRec = recordWindow()
                self.newRec.ui.setWindowModality(QtCore.Qt.ApplicationModal)
                self.newRec.ui.cbx_userList.setModel(self.userProxyModel)
                self.newRec.ui.cbx_userList.setCurrentIndex(self.ui.cbx_user.currentIndex())
                self.newRec.ui.but_OK.clicked.connect(self.editRecord) # connect recordWindow.but_OK to mainWindow.editRecord
                
                # apply parameters)
                if self.historyList:
                    self.newRec.ui.lst_history.addItems(self.historyList)
                self.newRec.ui.lst_history.setSortingEnabled(True)
                self.newRec.ui.lst_history.sortItems(QtCore.Qt.DescendingOrder)
                index = self.newRec.ui.cbx_userList.findText(recInfo.user, QtCore.Qt.MatchFixedString)
                if index >= 0:
                    self.newRec.ui.cbx_userList.setCurrentIndex(index)
                index = self.newRec.ui.cbx_type.findText(recInfo.Type, QtCore.Qt.MatchFixedString)
                if index >= 0:
                    self.newRec.ui.cbx_type.setCurrentIndex(index)
                self.newRec.ui.lin_nr.setText(str(recInfo.nr))
                self.newRec.ui.lin_project.setText(recInfo.project_path)
                self.newRec.ui.lin_render.setText(recInfo.rend_path)
                if recInfo.home:
                    self.newRec.ui.cbx_home.setChecked(True)
                else:
                    self.newRec.ui.cbx_home.setChecked(False)
                # apply parameters
                
                self.newRec.ui.show()
        self.filterRows()

    def goPlaces(self, ind):
        con = sqlite3.connect(dbPath)
        with con:
            cur = con.cursor()
            cur.execute('SELECT * FROM Projects WHERE ID=?', (recInfo.Id,))
            data = cur.fetchall()
            path = data[0][ind]
        if os.path.exists(path):
            subprocess.Popen('explorer "{0}"'.format(path))
    
    def editRecord(self):
        if self.newRec.result:
            self.addHistory()
            query = QtSql.QSqlQuery()
            query.prepare("UPDATE Projects SET User=:user, Nr=:nr, Project=:pr, Project_path=:pr_path, Type=:type, Rend_path=:rend_path, Home=:home WHERE ID=:id")
            query.bindValue(':user', recInfo.user)
            query.bindValue(':nr', recInfo.nr)
            query.bindValue(':pr', recInfo.project)
            query.bindValue(':pr_path', recInfo.project_path)
            query.bindValue(':type', recInfo.Type)
            query.bindValue(':rend_path', recInfo.rend_path)
            query.bindValue(':home', recInfo.home)
            query.bindValue(':id', recInfo.Id)
            query.exec_()
            self.updateChanges()

    def getRow(self):
        self.currRow = self.ui.tbl_View.selectionModel().currentIndex().row()
        recInfo.Id = self.ui.tbl_View.model().index(self.currRow, 0).data()
        recInfo.user = self.ui.tbl_View.model().index(self.currRow, 2).data()
        print(recInfo.Id, recInfo.user)
        self.displayMail()
        self.backUp()

    def closeEvent(self, event): # minimize instead of closing
        event.ignore()
        self.ui.setWindowState(QtCore.Qt.WindowMinimized) 
    
    def searchRecord(self):
        self.model.setFilterRegExp(QtCore.QRegExp(self.ui.lin_search.text(), QtCore.Qt.CaseInsensitive, QtCore.QRegExp.FixedString))
        #self.model.setFilterRegExp(QtCore.QRegExp(self.ui.lin_search.text())
        #self.model.setFilterKeyColumn(self.currColumn)
        self.filterRows()

    def openSettings(self): # OPEN SETTINGS
        self.setWin = settingsWindow()
        self.setWin.ui.setWindowModality(QtCore.Qt.ApplicationModal)
        self.setWin.setupDbView()
        
        def loadDB(): # load db and close settings window
            self.setWin.saveList() # save mail Inbox folders to file
            self.setWin.ui.close()
            self.getSettings()
            self.setupDbView()
            self.setFirstUser()
            self.recordMail()
            self.viewFilter()
            self.mailWatch()
        
        self.setWin.ui.but_OK.clicked.connect(lambda: loadDB())
        self.setWin.ui.show()

    def openRecord(self): # OPEN RECORD WINDOW
        self.newRec = recordWindow()
        self.newRec.ui.setWindowModality(QtCore.Qt.ApplicationModal)
        self.newRec.ui.but_OK.clicked.connect(self.makeRecord) # connect recordWindow.but_OK to mainWindow.makeRecord
        self.newRec.ui.cbx_userList.setModel(self.userProxyModel)
        self.newRec.ui.cbx_userList.setCurrentIndex(self.ui.cbx_user.currentIndex())
        self.newRec.ui.cbx_userList.setEnabled(self.ui.chb_Global.isChecked()) # disable if Global view is on
        self.newRec.ui.show()

    def makeRecord(self):
        if self.newRec.result:
            recInfo.status = 'In Progress'
            recInfo.props = time.strftime('%Y-%m-%d') + ' - ' + recInfo.status + ' - ' + recInfo.user + ','
            info = [
                    recInfo.props,          # 1
                    recInfo.user,           # 2
                    recInfo.priority,       # 3
                    recInfo.project,        # 4
                    recInfo.project_path,   # 5
                    recInfo.nr,             # 6
                    recInfo.Type,           # 7
                    recInfo.cad,            # 8
                    recInfo._3d2d,          # 9
                    recInfo.rend,           # 10
                    recInfo.rend_path,      # 11
                    recInfo.post,           # 12
                    recInfo.rev,            # 13
                    recInfo.notes,          # 15
                    recInfo.status,         # 16
                    recInfo.home            # 17
                    ]
            query = QtSql.QSqlQuery()
            query.prepare("INSERT INTO Projects (Props,User,Priority,Project,Project_path,Nr,Type,CAD,DD,Rend,Rend_path,Post,Rev,Notes,Status,Home) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)")
            for num, val in enumerate(info):
                query.bindValue(num, val)
            query.exec_()
            self.updateChanges()
    
    def delRecord(self):
        table = self.ui.tbl_View
        if table.currentIndex().row() != -1:
            reply = QMessageBox.question(self, 'Sure?',"Delete selected record?", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.Yes:
                if recInfo.Id == table.model().index(self.currRow, 0).data():
                    self.model.removeRow(table.currentIndex().row())
                    self.updateChanges()
            else:
                pass 

    def resizeTable(self):
        table = self.ui.tbl_View
        w = 45
        for i in (0,1,3,8,9,10,12,13,17): # fixed columns
            table.setColumnWidth(i, w)
            table.horizontalHeader().setSectionResizeMode(i, QHeaderView.Fixed)
        for i in (2,6,7,14): # fixed, wider columns
            table.setColumnWidth(i, 70)
            table.horizontalHeader().setSectionResizeMode(i, QHeaderView.Fixed)
        for i in (4,15): # stretch columns
            table.horizontalHeader().setSectionResizeMode(i, QHeaderView.Stretch)
        for i in range(table.model().columnCount()): # hide columns
            if i not in (3,4,6,7,8,9,10,12,13,14,15,17): #17 Home
                table.setColumnHidden(i,True)

    def viewFilter(self):
        table = self.ui.tbl_View
        table.setColumnHidden(2,not self.ui.chb_Global.isChecked())
        #table.setColumnHidden(16,not self.ui.chb_Global.isChecked())
        self.ui.cbx_user.setEnabled(not self.ui.chb_Global.isChecked())
        if not self.ui.chb_Global.isChecked():
            recInfo.user = self.ui.cbx_user.currentText()
        self.filterRows()
        self.displayMail()
        
    def filterRows(self):
        table = self.ui.tbl_View
        if table.model():
            status = self.ui.cbx_status.currentText()
            if not self.ui.chb_Global.isChecked():
                for i in range(table.model().rowCount()):
                    if table.model().index(i, 16).data() != status:
                        table.setRowHidden(i, True)
                    elif table.model().index(i, 2).data() != self.ui.cbx_user.currentText():
                        table.setRowHidden(i, True)
                    else:
                        table.setRowHidden(i, False)
            else:
                for i in range(table.model().rowCount()):
                    if table.model().index(i, 16).data() != status:
                        table.setRowHidden(i, True)
                    else:
                        table.setRowHidden(i, False)

    def setupDbView(self):
        if os.path.exists(dbPath):
            self.db = QtSql.QSqlDatabase.addDatabase('QSQLITE')
            self.db.setDatabaseName(dbPath)
            self.initializeModel()
            table = self.ui.tbl_View
            table.setModel(self.model)
            if table.model().headerData(17, QtCore.Qt.Horizontal) == 'Home':
                table.horizontalHeader().moveSection(17,15) # move the Home column
            table.setSortingEnabled(True)
            self.resizeTable()
            title = (appName + '[ ' + currUser + ' ]')
            self.ui.setWindowTitle(title)
            self.dbWatcher = QtCore.QFileSystemWatcher() # watch for file changes
            self.dbWatcher.addPath(os.path.dirname(dbPath))
            self.dbWatcher.directoryChanged.connect(self.updateChanges)
            table.selectionModel().selectionChanged.connect(self.getRow) # get the row index
    
    def updateChanges(self):
        self.remoteControl()
        self.sourceModel.select()
        while self.sourceModel.canFetchMore():
            self.sourceModel.fetchMore()
        self.filterRows()
        self.displayMail()
    
    class MySqlModel(QtSql.QSqlTableModel): # centre the cell data, disable some columns and set colours
        
        def data(self, index, role=QtCore.Qt.DisplayRole):
            if not index.isValid():
                return QtCore.QVariant()
            
            elif role == QtCore.Qt.TextAlignmentRole:
                if index.column() != 4: # do not centre 'project' column
                    return QtCore.Qt.AlignCenter
            
            elif role == QtCore.Qt.BackgroundRole:
                column13_data = index.sibling(index.row(), 13).data() # 'revision' column colour
                column17_data = index.sibling(index.row(), 17).data() # 'home' column colour
                if index.column() == 13:
                    if column13_data == 0:
                        return QtCore.QVariant(QtGui.QBrush(QtGui.QColor(0, 255, 0, 60))) # green
                    if column13_data > 0 and column13_data < 5:
                        return QtCore.QVariant(QtGui.QBrush(QtGui.QColor(255, 255, 0, 90))) # yellow
                    if column13_data > 4:
                        return QtCore.QVariant(QtGui.QBrush(QtGui.QColor(255, 0, 0, 60))) # red
                if index.column() != 13: # don't colour revision column
                    if column17_data:
                        return QtCore.QVariant(QtGui.QBrush(QtGui.QColor(QtCore.Qt.lightGray)))
                    
            return QtSql.QSqlTableModel.data(self,index,role)
        
        def flags(self, index):
            if index.column() not in (0,1,2,4,5,7,11,14,16): # make cells NOT EDITABLE
                return QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEditable
            else:
                return QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable
            
            
    class MySortFilterProxyModel(QtCore.QSortFilterProxyModel): # custom filter model
            
        def filterAcceptsRow(self, row, parent):
            model = self.sourceModel()
            regEx = self.filterRegExp()
            ind2 = model.index(row, 2, parent) # user
            ind4 = model.index(row, 4, parent) # project
            ind6 = model.index(row, 6, parent) # number
            ind15 = model.index(row, 15, parent) # comments
            return  regEx.indexIn(model.data(ind2)) >= 0 or \
                    regEx.indexIn(model.data(ind4)) >= 0 or \
                    regEx.indexIn(str(model.data(ind6))) >= 0 or \
                    regEx.indexIn(model.data(ind15)) >= 0
    

    def initializeModel(self):
        self.sourceModel = self.MySqlModel()
        self.sourceModel.setTable('Projects')
        self.sourceModel.setEditStrategy(QtSql.QSqlTableModel.OnFieldChange)
        
        #self.sourceModel.select()
        #for i in range(self.sourceModel.rowCount()):
            #self.sourceModel.index(i,0).item() # QtCore.Qt.ItemIsSelectable
            #print(table.setItemDelegateForColumn(0))
        
        self.userModel = QtSql.QSqlTableModel()
        self.userModel.setTable('Users')
        self.userProxyModel = QtCore.QSortFilterProxyModel()
        self.userProxyModel.setSourceModel(self.userModel)
        self.userModel.select()
        while self.userModel.canFetchMore():
            self.userModel.fetchMore()
        self.ui.cbx_user.setModel(self.userProxyModel)
        self.ui.cbx_user.setModelColumn(self.userModel.fieldIndex('User'))
        self.ui.cbx_user.model().sort(0)
        
        self.model = self.MySortFilterProxyModel() # custom proxy model
        self.model.setSourceModel(self.sourceModel)
        self.updateChanges()
        
##        self.model.setHeaderData(0, QtCore.Qt.Horizontal, "ID")
##        self.model.setHeaderData(1, QtCore.Qt.Horizontal, "Props")
##        self.model.setHeaderData(2, QtCore.Qt.Horizontal, "User")
        self.model.setHeaderData(3, QtCore.Qt.Horizontal, "Pr-ty")
##        self.model.setHeaderData(4, QtCore.Qt.Horizontal, "Project")
##        self.model.setHeaderData(5, QtCore.Qt.Horizontal, "Project_path")
        self.model.setHeaderData(6, QtCore.Qt.Horizontal, "Ref no")
##        self.model.setHeaderData(7, QtCore.Qt.Horizontal, "Type")
##        self.model.setHeaderData(8, QtCore.Qt.Horizontal, "CAD")
        self.model.setHeaderData(9, QtCore.Qt.Horizontal, "3D-2D")
##        self.model.setHeaderData(10, QtCore.Qt.Horizontal, "Rend")
##        self.model.setHeaderData(11, QtCore.Qt.Horizontal, "Rend_path")
##        self.model.setHeaderData(12, QtCore.Qt.Horizontal, "Post")
##        self.model.setHeaderData(13, QtCore.Qt.Horizontal, "Rev")
##        self.model.setHeaderData(14, QtCore.Qt.Horizontal, "Created")
##        self.model.setHeaderData(15, QtCore.Qt.Horizontal, "Notes")
##        self.model.setHeaderData(16, QtCore.Qt.Horizontal, "Status")

        
class recordWindow(QDialog):

    def __init__(self):
        QDialog.__init__(self)
        self.ui = loadUi('proMan_record.ui')
        self.result = False

        self.ui.lin_project.textChanged.connect(self.projectLine)
        self.ui.but_OK.clicked.connect(self.addRecord)
        self.ui.but_Cancel.clicked.connect(self.Nope)
        self.ui.lin_nr.setValidator(QtGui.QIntValidator(0, 99999, self))

    def projectLine(self):
        if self.ui.lin_project.text() != '':
            text = self.ui.lin_project.text()
            textLower = text.lower()
            if textLower.startswith(r'\\vr\c\clients'):
                text = text[15:]
                tList = text.split(os.sep)
                text = ' \ '.join(tList)
            self.ui.lab_project.setText(text)
            self.ui.but_OK.setEnabled(True)
        else:
            self.ui.lab_project.setText('...')
            self.ui.but_OK.setEnabled(False)

    def addRecord(self):
        if os.path.exists(self.ui.lin_project.text()):
            # filling out record info
            recInfo.user = self.ui.cbx_userList.currentText()
            recInfo.nr = self.ui.lin_nr.text()
            recInfo.project = self.ui.lab_project.text()
            recInfo.project_path = self.ui.lin_project.text()
            recInfo.Type = self.ui.cbx_type.currentText()
            recInfo.rend_path = self.ui.lin_render.text()
            if self.ui.cbx_home.isChecked():
                if not recInfo.home:
                    recInfo.home = 'y'
            else:
                recInfo.home = ''
            # filling out record info
            self.ui.close()
            self.result = True
        else:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText('\nThis project directory does not exist!      ')
            msg.setWindowTitle('Nope!')
            msg.exec_()
            self.result = False

    def Nope(self):
        self.ui.close()

                        
class settingsWindow(QDialog):

    def __init__(self):
        QDialog.__init__(self)
        self.ui = loadUi('proMan_settings.ui')

        self.ui.but_Browse.clicked.connect(self.browseDB)
        self.ui.but_newDB.clicked.connect(self.openPass)
        self.ui.but_createUser.clicked.connect(self.makeRecord)
        self.ui.but_deleteUser.clicked.connect(self.delRecord)
        self.ui.but_addPath.clicked.connect(self.addPath)
        self.ui.but_deletePath.clicked.connect(self.delPath)
        
        self.currRow = 0
        self.fillList()
        
        if dbPath:
            self.ui.lin_dbLocation.setText(dbPath)
            self.ui.but_createUser.setEnabled(True)
            self.ui.but_deleteUser.setEnabled(True)
            self.ui.but_OK.setEnabled(True)

    def fillList(self):
        self.ui.lst_Mail.clear()
        if mailList:
            for item in mailList:
                if item:
                    self.ui.lst_Mail.addItem(item)
            
    def saveList(self):
        folders = ''
        for i in range(self.ui.lst_Mail.count()):
            folders += self.ui.lst_Mail.item(i).text() + ','
        m = open(mailIni, 'w')
        m.write(folders)
        m.close()

    def addPath(self):
        inbox = str(pathlib.Path(root).parents[0]) + '\\Local\\Microsoft\\Windows Live Mail\\'
        bm = QFileDialog.getExistingDirectory(self, 'Browse for mail Inbox', inbox)
        if bm:
            self.ui.lst_Mail.addItem(bm)
    
    def delPath(self):
        row = self.ui.lst_Mail.currentRow()
        if row > -1:
            self.ui.lst_Mail.takeItem(row)

    def browseDB(self):
        db = QFileDialog.getOpenFileNames(self, 'Browse for database file', '*.db')
        if db[0]:
            global dbPath
            dbPath = db[0][0]
            self.ui.but_OK.setEnabled(True)
            self.ui.lin_dbLocation.setText(dbPath)
            if not os.path.exists(setDir): os.makedirs(setDir)
            d = open(dbIni, 'w')
            d.write(dbPath)
            d.close()
    
    def openPass(self):
        self.passWin = passwordWindow()
        self.passWin.ui.setWindowModality(QtCore.Qt.ApplicationModal)
        self.passWin.ui.but_OK.clicked.connect(self.checkPass)
        self.passWin.ui.show()
        
    def checkPass(self):
        if self.passWin.ui.lin_pass.text() == 'password':
            self.freshDB()
            self.passWin.ui.close()
        else:
            self.passWin.ui.close()
        
    def freshDB(self):
        ndb = QFileDialog.getSaveFileName(self, 'Choose save location', '*.db')
        if ndb[0]:
            con = sqlite3.connect(ndb[0])
            with con:
                cur = con.cursor()
                cur.execute("CREATE TABLE Projects(ID INTEGER PRIMARY KEY UNIQUE, Props TEXT, User TEXT, Priority INT, \
                    Project TEXT, Project_path TEXT, Nr INT, Type TEXT, CAD TEXT, DD TEXT, Rend TEXT, \
                    Rend_path TEXT, Post TEXT, Rev INT, Created DATE DEFAULT CURRENT_DATE, Notes TEXT, Status TEXT, Home TEXT, \
                    Info_A TEXT, Info_B TEXT, Info_C TEXT, Info_D TEXT, Info_E TEXT, Info_F TEXT, Info_G TEXT, Info_H TEXT)")
                cur.execute("CREATE TABLE Users(User TEXT NOT NULL UNIQUE, Mail TEXT)")
            global dbPath
            dbPath = ndb[0]
            d = open(dbIni, 'w')
            d.write(dbPath)
            d.close()
            self.ui.lin_dbLocation.setText(dbPath)
            self.ui.but_createUser.setEnabled(True)
            self.ui.but_deleteUser.setEnabled(True)
            self.ui.but_OK.setEnabled(True)
            
    def getRow(self):
        self.currRow = self.ui.tbl_Users.selectionModel().currentIndex().row()
            
    def makeRecord(self):
        if self.ui.lin_newUser.text():
            val = self.ui.lin_newUser.text()
            query = QtSql.QSqlQuery()
            query.prepare("INSERT INTO Users (User) VALUES (?)")
            query.bindValue(0, val)
            res = query.exec_()
            if res:
                self.sourceModel.select()
                while self.sourceModel.canFetchMore():
                    self.sourceModel.fetchMore()
                self.ui.lin_newUser.setText('')
            else:
                msg = QMessageBox()
                msg.setIcon(QMessageBox.Critical)
                msg.setText('\nThis user name already exists!      ')
                msg.setWindowTitle('Bad user name')
                msg.exec_()
    
    def delRecord(self):
        table = self.ui.tbl_Users
        if table.currentIndex().row() != -1:
            reply = QMessageBox.question(self, 'Sure?',"Delete selected user?", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.Yes:
                self.model.removeRow(table.currentIndex().row())
                self.sourceModel.select()
                while self.sourceModel.canFetchMore():
                    self.sourceModel.fetchMore()
            else:
                pass         
    
    def resizeUserTable(self):
        table = self.ui.tbl_Users
        for i in range(table.model().columnCount()): # hide columns
            if not i == 0:
                table.setColumnHidden(i,True)
    
    def setupDbView(self):
        if os.path.exists(dbPath):
            self.db = QtSql.QSqlDatabase.addDatabase('QSQLITE')
            self.db.setDatabaseName(dbPath)
            self.initializeModel()
            table = self.ui.tbl_Users
            table.setModel(self.model)
            table.sortByColumn(0, QtCore.Qt.AscendingOrder)
            table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
            self.resizeUserTable()
            self.dbWatcher = QtCore.QFileSystemWatcher() # watch for file changes
            self.dbWatcher.addPath(os.path.dirname(dbPath))
            self.dbWatcher.directoryChanged.connect(lambda: self.sourceModel.select())
            table.selectionModel().selectionChanged.connect(self.getRow) # get the row index
        
    def initializeModel(self):
        self.sourceModel = QtSql.QSqlTableModel()
        self.sourceModel.setTable('Users')
        self.sourceModel.setEditStrategy(QtSql.QSqlTableModel.OnFieldChange)
        self.model = QtCore.QSortFilterProxyModel()
        self.model.setSourceModel(self.sourceModel)
        self.sourceModel.select()
        while self.sourceModel.canFetchMore():
            self.sourceModel.fetchMore()
        
        
class passwordWindow(QDialog):
    def __init__(self):
        QDialog.__init__(self)
        self.ui = loadUi('proMan_pass.ui')
        

if __name__ == '__main__':

    app = QApplication(sys.argv)
    window = mainWindow()
    sys.exit(app.exec_())
