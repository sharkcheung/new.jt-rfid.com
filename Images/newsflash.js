var swf_width=1002 
var swf_height=250 
var files='/images/banner1.jpg|/images/banner2.jpg|/images/banner3.jpg' 
var links='' 
var texts='' 
document.write('<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" width="'+ swf_width +'" height="'+ swf_height +'">');
document.write('<param name="movie" value="/images/flash.swf"><param name="quality" value="high">');
document.write('<param name="menu" value="false"><param name=wmode value="opaque">');
document.write('<param name="FlashVars" value="bcastr_file='+files+'&bcastr_link='+links+'&bcastr_title='+texts+'&bcastr_config=0xFF0000:文字颜色|2:文字位置|0x555555:文字背景颜色|90:文字背景透明度|0xffffff:按键文字颜色|0x1480B8:按键默认颜色|0x555555:按键当前颜色|5:自动播放时间(秒)|3:图片过渡效果|1:是否显示按钮|_top:打开窗口">');
document.write('<embed src="/images/flash.swf" wmode="opaque" FlashVars="bcastr_file='+files+'&bcastr_link='+links+'&bcastr_title='+texts+'& menu="false" quality="high" width="'+ swf_width +'" height="'+ swf_height +'&bcastr_config=0x002567:文字颜色|1:文字位置|0x1480B8:文字背景颜色|70:文字背景透明度|0xffffff:按键文字颜色|0xE00082:按键默认颜色|0x1480B8:按键当前颜色|5:自动播放时间(秒)|3:图片过渡效果|1:是否显示按钮|_top:打开窗口" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />'); 
document.write('</object>'); 

/* 高级设置 默认参数字符串 0xffffff:文字颜色|1:文字位置|0xff6600:文字背景颜色|60:文字背景透明度|0xffffff:按键文字颜色|0xff6600:按键默认颜色|0x000033:按键当前颜色|8:自动播放时间(秒)|3:图片过渡效果|1:是否显示按钮|_blank:打开窗口 颜色都以0x开始16进制数字表示 文字颜色：题目文字的颜色 文字位置：0表示题目文字在顶端，1表示文字在底部，2表示文字在顶端 文字背景透明度：0-100值，0表示全部透明 按键文字颜色：按键数字颜色 按键默认颜色：按键默认的颜色 按键当前颜色：当前图片按键颜色 自动播放时间：单位是秒 图片过渡效果：0，表示亮度过渡，1表示透明度过渡，2表示模糊过渡，3表示运动模糊过渡 是否显示按钮：0，表示隐藏按键部分，更适合做广告挑轮换 影片自动播放参数：0表示不自动播放，1表示自动播放 影片连续播放参数：0表示不连续播放，1表示连续循环播 默认音量参数 ：0-100 的数值，设置影片开始默认音量大小 控制栏位置参数 ：0表示在影片上浮动显示，1表示在影片下方显示 控制栏显示参数 ：0表示不显示；1表示一直显示；2表示鼠标悬停时显示；3表示开始不显示，鼠标悬停后显示 打开窗口：_blank表示新窗口打开。_self表示在当前窗口打开 */
