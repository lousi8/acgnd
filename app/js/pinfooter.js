$(function(){
	$("img.lazy").lazyload({		
		load:function(){
			$('#container').BlocksIt();
		}
	});	
	$(window).scroll(function(){
			// 当滚动到最底部以上50像素时， 加载新内容
		if ($(document).height() - $(this).scrollTop() - $(this).height()<50){
			$('#container').append($("#test").html());		
			$('#container').BlocksIt();
			$("img.lazy").lazyload();
		}
	});
	
	//window resize
	$(window).resize(function() {
	$('#container').BlocksIt();
	});

	//window onorientationchange
   $(window).on("orientationchange",function(event){
  //alert("方向是：" + event.orientation);
  $('#container').BlocksIt();
});

});
var _bdhmProtocol = (("https:" == document.location.protocol) ? " https://" : " http://");
document.write(unescape("%3Cscript src='" + _bdhmProtocol + "hm.baidu.com/h.js%3F007d7ca43e962ab9055d7507e1299a25' type='text/javascript'%3E%3C/script%3E"));
