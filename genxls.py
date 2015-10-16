#!/usr/bin/python
#!_*_ coding:utf-8 _*_
import sys
#reload(sys)
#sys.setdefaultencoding('utf-8')
import getopt
import xlrd
import xlwt
class myexception(Exception):pass
class maker(object):
	def __init__(self,startnum=1,endnum=50000,contentstart=1,percount=500,wfilename='wfile',sheetname='a sheet'):
		self.startnum=startnum
		self.connumberlen=len(str(contentstart))
		self.endnum=endnum
		self.percount=percount
		self.sheetname=sheetname
		self.contentstart=int(contentstart)
		self.contentend=contentstart+endnum-startnum
		self.wfilename=wfilename
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
				w.save(self.wfilename+str(x+1)+'.xls')
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
	startnum=None
	endnum=None
	constart=None
	percount=None
	wfilename=None
	rfilename=None
	try:
		opts,args=getopt.getopt(sys.argv[1:],'hs:e:c:p:w:r:',['help','startnum=','endnum=','constart=','percount=','wfilename=','rfilename='])
		for k,v in opts:
			if k in ('-h','--help'):
				showhelp()
				sys.exit(0)
			elif k in ('-r','--rfilename'):
				rfilename=v
			elif k in ('-s','--startnum'):
				startnum=int(v)
			elif k in ('-e','--endnum'):
				endnum=int(v)
			elif k in ('-c','--constart'):
				constart=v
			elif k in ('-p','--percount'):
				percount=int(v)
			elif k in ('-w','--wfilename'):
				wfilename=v
		if rfilename is not None:
			m=maker()
			m.startread(rfilename)
			print 'Read Finished'
		elif None in (startnum,endnum,constart,percount,wfilename):
			print 'Must enter the all args to write!'
			showhelp()
			sys.exit(1)
		else:
			m=maker(startnum,endnum,constart,percount,wfilename)
			status,info=m.startwrite()
			print info
	except getopt.GetoptError,e:
		print str(e)
		showhelp()
		sys.exit(1)
	except ValueError,e:
		print str(e)
		showhelp()
		sys.exit(1)
