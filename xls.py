#!_*_ coding:utf-8 _*_
import json
import genxls
def index(req,**args):
	return _get_html()

def handler(req,**args):

	startid=int(args['startid'])
	endid=int(args['endid'])
	startphonenum=int(args['startphonenum'])
	savefilepath=args['savefilepath']
	savecountsperfile=int(args['savecountsperfile'])
	status,info=genxls.maker(startnum=startid,endnum=endid,contentstart=startphonenum,percount=savecountsperfile,wfilename=savefilepath).startwrite()
	return json.dumps({'resultstatus':status,'info':info})
def _get_html():
	return '''
<html>
<head>
<script type="text/javascript" src="static/jquery.js"></script>
<script type="text/javascript" src="static/my.js"></script>
</script>
</head>
<body>
<div id="statusbar">

</div>
<div>
<form action="" method="post" name="form" id="form">
<table border="1">
<caption align="top">自动生存XLS文件</caption>
<tr>
<td>StartID:</td>
<td><input type="text" name="startid"/></td>
</tr>
<tr>
<td>EndID:</td>
<td><input type="text" name="endid"/></td>
</tr>
<tr>
<td>StartPhoneNum:</td>
<td><input type="text" name="startphonenum"/></td>
</tr>
<tr>
<td>SaveCountsPerFile:</td>
<td><input type="text" name="savecountsperfile"/></td>
</tr>
<tr>
<td>SaveFilePath:</td>
<td><input type="text" name="savefilepath"/></td>
</tr>
<tr>
<td align="right">
<input type="button" name="subit" id="submit" value="OK" style="width:100px;"/>
</td>
<td>
<input type="reset" name="reset" value="clear" style="width:100px;"/>
</td>
</tr>
</table>
</form>
</div>
</body>
</html>
'''
