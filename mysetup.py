from distutils.core import setup
import py2exe
setup(
	windows=[{'script':'wingui.py'}],
	options={'py2exe':{'packages':['wx.lib.pubsub',
	]}}
	)
