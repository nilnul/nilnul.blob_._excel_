<%@ Control Language="C#" AutoEventWireup="true"  %>
<p>
	发信人: SlANmASTer (渴望美女青睐 之 我爱工科女), 信区: WebDev<br />
	标 题: web导出excel文件的几种方法zz<br />
	发信站: 水木社区 (Sat Apr 20 15:50:18 2013), 站内<br />
	<br />
	web导出excel文件的几种方法<br />
	<br />
	<br />
	<br />
	KimmKing<br />
	<br />
	kimmking@163.com<br />
	<br />
	2009年9月4日10:19:09<br />
	<br />
	<br />
	<br />
	总的来说，两种方法：服务器端生成和浏览器端生成。<br />
	<br />
	<br />
	<br />
	服务器端生成就是：根据用户请求，获取相应的数据，使用poi/jxl, jacob/jawin+excel,或是用数据拼html的table或是cvs纯文本的数据格式等。然后按.xls或是.cvs格式的文件的形式返回给用户，指定Content-Type:application/vnd.ms-excel ,浏览器就会提示要下载的文件是excel文件。<br />
	<br />
	poi/jxl, jacob/jawin生成的是excel的biff格式。html/csv的是文本格式，不另存为excel文件，很多excel功能是用不了的。jacob/jawin需要服务器端是windows系统，且安装了excel2000以上版本。poi/jxl和html/csv方式的话，服务器端可以跨平台。<br />
	<br />
	<br />
	<br />
	浏览器端生成excel文件还没有特别完善的方案，这是因为js无法处理二进制。大概有以下几个方案，各有利弊。<br />
	<br />
	1. activex方式：使用js/vbs调用excel对象，<a href="http://setting.iteye.com/blog/219302" target="_blank">http://setting.iteye.com/blog/219302</a>，有个extjs的gridpanel导出为excel的例子。 （ie+excel）<br />
	<br />
	2. ie命令方式：将html或是csv输出到open的window，然后使用execCommand的saveas命令，存为csv或xls。 (ie6 only)<br />
	<br />
	3. 服务器端中转方式：将html的table或是拼接的csv传到服务器端，服务器端再按照Content-Type:application/vnd.ms-excel返回，浏览器就会按excel方式处理。与服务器端拼接相比，少了一次取数操作。 (all)<br />
	<br />
	4. data协议方式：对于支持data协议的浏览器，可以将html或是csv先用js base64处理，然后前缀data:application/vnd.ms-excel;base64,，即可使浏览器将其中的数据当做excel来处理，浏览器将提示下载或打开excel文件,可惜的是ie不支持。extjs的官网有一个grid的plugin，实现导出xhtml格式的伪excel文件，就是这么做的。 (except IE)<br />
	<br />
	<br />
	<br />
	浏览器端只有第一种方案导出的是真正的biff格式的excel文件，其他方式都是文本格式。activex方式只能在windows平台的ie浏览器使用，而且需要降低ie的安全性，所以应用比较有限。复杂的excel文件，还是在服务器端用poi/jxl生成excel比较好。如果浏览器固定位ie6，浏览器端方式2是最好的方案。如果要降低服务器端cpu的计算压力，客户端方案3可行，而且跨平台（比poi/jxl方式少了取数和生成二进制文件）。如果是非ie浏览器，方案4也不失为一种好方法。<br />
	<br />
	<br />
	<br />
	<br />
	<br />
	<br />
	ps: 还有一个方案，就是让安装了ie和excel的用户在网页上右键，点击“导出到 Microsoft Excel”，然后可以选择要导出的table区域，点“导入”按钮，完成导入。<br />
	<br />
	--<br />
	我本将心比明月，奈何明月照沟渠<br />
</p>

