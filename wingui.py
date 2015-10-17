#!/usr/bin/python
#!_*_ coding:utf-8 _*_
import wx
#from wx.lib.pubsub import Publisher
from wx.lib.pubsub import setuparg1
from wx.lib.pubsub import pub as Publisher
import threading
import sys
import xlwt
import xlrd
import os
class myexception(Exception):pass

class mygui(wx.Frame):
	def __init__(self):
		wx.Frame.__init__(self,None,title='XLS Maker')
		self.SetSizeHintsSz((300,280),(300,280))
		self.panel=wx.Panel(self)
		self.Bind(wx.EVT_CLOSE,self.closeaction)
		self.startid=wx.TextCtrl(self.panel)
		self.endid=wx.TextCtrl(self.panel)
		self.startphonenum=wx.TextCtrl(self.panel)
		self.savecountsperfile=wx.TextCtrl(self.panel)
		self.savefilepath=wx.TextCtrl(self.panel)
		self.countsperfolder=wx.TextCtrl(self.panel)
		self.okbutton=wx.Button(self.panel,label='ok')
		self.okbutton.Bind(wx.EVT_BUTTON,self.okaction)
		self.exitbutton=wx.Button(self.panel,label='exit')
		self.exitbutton.Bind(wx.EVT_BUTTON,self.exitaction)
		self.hbox1=wx.BoxSizer()
		self.hbox1.Add(wx.StaticText(self.panel,label='StartID:'),proportion=1,flag=wx.EXPAND,border=0)
		self.hbox1.Add(self.startid,proportion=1,flag=wx.EXPAND,border=0)
		self.hbox2=wx.BoxSizer()
		self.hbox2.Add(wx.StaticText(self.panel,label='EndID:'),proportion=1,flag=wx.EXPAND,border=0)
		self.hbox2.Add(self.endid,proportion=1,flag=wx.EXPAND,border=0)
		self.hbox3=wx.BoxSizer()
		self.hbox3.Add(wx.StaticText(self.panel,label='StartPhoneNum:'),proportion=1,flag=wx.EXPAND,border=0)
		self.hbox3.Add(self.startphonenum,proportion=1,flag=wx.EXPAND,border=0)
		self.hbox4=wx.BoxSizer()
		self.hbox4.Add(wx.StaticText(self.panel,label='SaveCountsPerFile:'),proportion=1,flag=wx.EXPAND,border=0)
		self.hbox4.Add(self.savecountsperfile,proportion=1,flag=wx.EXPAND,border=0)
		self.hbox5=wx.BoxSizer()
		self.hbox5.Add(wx.StaticText(self.panel,label='SaveFileName:'),proportion=1,flag=wx.EXPAND,border=0)
		self.hbox5.Add(self.savefilepath,proportion=1,flag=wx.EXPAND,border=0)
		self.hbox7=wx.BoxSizer()
		self.hbox7.Add(wx.StaticText(self.panel,label='CountsPerFolder:'),proportion=1,flag=wx.EXPAND,border=0)
		self.hbox7.Add(self.countsperfolder,proportion=1,flag=wx.EXPAND,border=0)
		self.hbox6=wx.BoxSizer()
		self.hbox6.Add(self.okbutton,proportion=1,flag=wx.EXPAND,border=0)
		self.hbox6.Add(self.exitbutton,proportion=1,flag=wx.EXPAND,border=0)
		self.vbox=wx.BoxSizer(wx.VERTICAL)
		self.vbox.Add(self.hbox1,border=5)
		self.vbox.Add(self.hbox2,border=5)
		self.vbox.Add(self.hbox3,border=5)
		self.vbox.Add(self.hbox4,border=5)
		self.vbox.Add(self.hbox5,border=5)
		self.vbox.Add(self.hbox7,border=5)
		self.vbox.Add(self.hbox6,border=5)
		self.panel.SetSizer(self.vbox)
		Publisher.subscribe(self.makeresult,'makeresult')
	def closeaction(self,evt):
		print 'close'
	def okaction(self,evt):
		try:
			startid=int(self.startid.GetValue().strip())
			endid=int(self.endid.GetValue().strip())
			startphonenum=self.startphonenum.GetValue().strip()
			savecountsperfile=int(self.savecountsperfile.GetValue().strip())
			countsperfolder=int(self.countsperfolder.GetValue().strip())
			savefilepath=self.savefilepath.GetValue().strip()
			mt=makethread(startid,endid,startphonenum,savecountsperfile,savefilepath,countsperfolder)
			mt.start()
		except Exception,e:
			dlg=wx.MessageDialog(self.panel,str(e),caption='Error',style=wx.OK)
			dlg.ShowModal()
			dlg.Destroy()

	def exitaction(self,evt):
		sys.exit(0)

	def makeresult(self,result):
		dlg=wx.MessageDialog(self.panel,result.data,caption='Result',style=wx.OK)
		dlg.ShowModal()
		dlg.Destroy()
class makethread(threading.Thread):
	def __init__(self,startid,endid,startphonenum,savecounts,savepath,countsperfolder):
		threading.Thread.__init__(self)
		self.startid=startid
		self.endid=endid
		self.startphonenum=startphonenum
		self.savecounts=savecounts
		self.savepath=savepath
		self.countsperfolder=countsperfolder
	def run(self):
		mkr=maker(startnum=self.startid,endnum=self.endid,contentstart=self.startphonenum,percount=self.savecounts,wfilename=self.savepath,countsperfolder=self.countsperfolder)
		status,info=mkr.startwrite()
		wx.CallAfter(self.postdata,info)
	def postdata(self,info):
		Publisher.sendMessage('makeresult',info)
class maker(object):
	def __init__(self,startnum=1,endnum=50000,contentstart=1,percount=500,wfilename='wfile',sheetname='a sheet',countsperfolder=10):
		self.startnum=startnum
		self.connumberlen=len(str(contentstart))
		self.endnum=endnum
		self.percount=percount
		self.sheetname=sheetname
		self.contentstart=int(contentstart)
		self.contentend=self.contentstart+endnum-startnum
		self.wfilename=wfilename
		self.countsperfolder=countsperfolder
	def checknumber(self,startnum,endnum):
#		if startnum>endnum:
#			raise myexception('invalid input number')
		self.filenumbers=(endnum-startnum)//self.percount
		self.numberlen=len(str(endnum))
#		self.connumberlen=len(str(self.contentend))
		if(endnum-startnum)%self.percount!=0:
			self.filenumbers+=1
		self.status=1
	def startwrite(self):
		try:
			self.checknumber(self.startnum,self.endnum)
			headers=(u'姓名',u'手机',u'宅话',u'Email',u'地址',u'单位',u'职务')
			for x in xrange(self.filenumbers):
				w=xlwt.Workbook()
				ws=w.add_sheet(self.sheetname)
				starts=self.startnum+self.percount*x
				constarts=self.contentstart+self.percount*x
				if x==self.filenumbers-1:
					left=(self.endnum-self.startnum)%self.percount
					if left:
						ends=starts+left+1
				else:
					ends=starts+self.percount
				interval=0
				for i,c in enumerate(headers):
					ws.write(0,i,c)
					ws.col(i).width=6000 if c==u'手机' else 3000
				for number in xrange(starts,ends):
					xline=(number-starts+1)
					#ws.write(xline,0,'%0{0}d'.format(self.numberlen)%(number))
					ws.write(xline,0,u'%d号'%(number))
					ws.write(xline,1,'%0{0}d'.format(self.connumberlen)%(constarts+interval))
					interval+=1
				dirpath='folder%d'%(x/self.countsperfolder+1)
				if x%self.countsperfolder==0:
					os.mkdir(os.path.join('result',dirpath))
				w.save(os.path.join('result',dirpath,self.wfilename+str(x+1)+'.xls'))
		except (Exception,myexception),e:
			return False,str(e)
		else:
			return True,'Write Succeed'
	def startread(self,filename):
#		try:
		bk=xlrd.open_workbook(filename)
		nsheets=bk.nsheets
		for n in xrange(nsheets):
			sh=bk.sheet_by_index(n)
			nrows=sh.nrows
			ncols=sh.ncols
			print '%s\tsheetname:%s\trows:%d\tcolumns:%d'%(filename,sh.name,nrows,ncols)
			print '\t'.join(['Column'+str(n+1) for n in xrange(ncols)])
			for r in xrange(nrows):
				tmp=''
				for c in xrange(ncols):
					tmp+='%s\t'%sh.cell_value(r,c)
				print '%s\n'%tmp.rstrip('\t')
#		except Exception,e:
#			pass
def showhelp():
	print '''Usage:
	-s or --startnum integer type
	-e or --endnum   integer type
	-c or --constart integer type
	-p or --percount integer type
	-w or --wfilename string type
	-r or --rfilename string type
	e.x.m:
		python test.py -r xlsfilename
		python test.py -s 10 -e 50000 -c 100860001 -p 500 -w /root/test11/w    (will generate some files which starts with 'w' in your wfilename location.first cell start from 10 and last cell is 50000,second cell is starting from 100860001.it will have 500 rows per file.)
		'''
if __name__=='__main__':
	app=wx.App()
	mg=mygui()
	mg.Show()
	app.MainLoop()
