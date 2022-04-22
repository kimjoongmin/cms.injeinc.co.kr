
var secTop = {};

$(function(){
	$("html, body").scrollTop(0);

	$("#section1").addClass("active");
	
	$("#slider ul").bxSlider({
		auto:true,
		autoControls:true,
		pager:true,
		pagerType:'short'
	});

	if($("#container .section").length > 0){
		var mainSection = $("#container .section");
		for(i=0; i<mainSection.length; i++){
			secTop[(i+1)] = mainSection.eq(i).offset().top;
		}
	}

	$("#sectionBtn li a").click(function(e){
		e.preventDefault();
		var no = $(this).parent().index() + 1;
		mainSectionControl.move(no);
	});

	$(".btn_gnbOpen").click(function(e){
		e.preventDefault();
		$("#gnb").addClass("active");
		$("#overlay").show();
	});
	$(".btn_gnbClose").click(function(e){
		e.preventDefault();
		$("#gnb").removeClass("active");
		$("#overlay").hide();
	});
	$("#gnb li a").click(function(e){
		if(!$(this).parent().hasClass("onlyPhone")){
			e.preventDefault();
			var no = $(this).parent().index() + 1;
			mainSectionControl.move(no);
			$("#gnb").removeClass("active");
			$("#overlay").hide();
		}
	});
	$(".btn_mobileZoom").click(function(e){
		e.preventDefault();
		if($(window).width() < 640){
			var targetUrl = $(this).children("img").attr("src");
			window.open("/zoom.html?url="+targetUrl);
		}
	});
});

$(window).resize(function(){
	if($("#container .section").length > 0){
		var mainSection = $("#container .section");
		for(i=0; i<mainSection.length; i++){
			secTop[(i+1)] = mainSection.eq(i).offset().top;
		}
	}
});

$(window).scroll(function(){
	var scrollTopPx = $(window).scrollTop() + ($(window).height()/2);
	mainSectionControl.scroll(scrollTopPx);
});

var mainSectionControl = {
	"down": function(){
		var scrollTopPx = $(window).scrollTop();
		if(scrollTopPx >= secTop[6]){
			$("html, body").stop().animate({"scrollTop": secTop[7]+"px"}, 300);

		}else if(scrollTopPx >= secTop[5]){
			$("html, body").stop().animate({"scrollTop": secTop[6]+"px"}, 300);

		}else if(scrollTopPx >= secTop[4]){
			$("html, body").stop().animate({"scrollTop": secTop[5]+"px"}, 300);

		}else if(scrollTopPx >= secTop[3]){
			$("html, body").stop().animate({"scrollTop": secTop[4]+"px"}, 300);

		}else if(scrollTopPx >= secTop[2]){
			$("html, body").stop().animate({"scrollTop": secTop[3]+"px"}, 300);

		}else if(scrollTopPx >= secTop[1]){
			$("html, body").stop().animate({"scrollTop": secTop[2]+"px"}, 300);

		}
	},
	"up": function(){
		var scrollTopPx = $(window).scrollTop();
		if(scrollTopPx <= secTop[1]){

		}else if(scrollTopPx <= secTop[2]){
			$("html, body").stop().animate({"scrollTop": secTop[1]+"px"}, 300);

		}else if(scrollTopPx <= secTop[3]){
			$("html, body").stop().animate({"scrollTop": secTop[2]+"px"}, 300);

		}else if(scrollTopPx <= secTop[4]){
			$("html, body").stop().animate({"scrollTop": secTop[3]+"px"}, 300);

		}else if(scrollTopPx <= secTop[5]){
			$("html, body").stop().animate({"scrollTop": secTop[4]+"px"}, 300);

		}else if(scrollTopPx <= secTop[6]){
			$("html, body").stop().animate({"scrollTop": secTop[5]+"px"}, 300);

		}else if(scrollTopPx <= secTop[7]){
			$("html, body").stop().animate({"scrollTop": secTop[6]+"px"}, 300);

		}
	},
	"move": function(no){
		var targetTop = $("#section"+no).offset().top;
		if($(window).width() < 1024){
			targetTop -= 60;
		}else{
			targetTop -= 100;
		}
		$("html, body").stop().animate({"scrollTop":targetTop+"px"}, 800);
	},
	"scroll": function(scrollTopPx){
		$("html").removeClass("scrolled_row1").removeClass("scrolled_row2").removeClass("scrolled_row3").removeClass("scrolled_row4").removeClass("scrolled_row5").removeClass("scrolled_row6").removeClass("scrolled_row7");
		if(scrollTopPx >= secTop[7]){
			$("html").addClass("scrolled_row7");
			$("#section7").addClass("active");
			$("#sectionBtn li").removeClass("active");
			$("#sectionBtn li").eq(6).addClass("active");

		}else if(scrollTopPx >= secTop[6]){
			$("html").addClass("scrolled_row6");
			$("#section6").addClass("active");
			$("#sectionBtn li").removeClass("active");
			$("#sectionBtn li").eq(5).addClass("active");

		}else if(scrollTopPx >= secTop[5]){
			$("html").addClass("scrolled_row5");
			$("#section5").addClass("active");
			$("#sectionBtn li").removeClass("active");
			$("#sectionBtn li").eq(4).addClass("active");

		}else if(scrollTopPx >= secTop[4]){
			$("html").addClass("scrolled_row4");
			$("#section4").addClass("active");
			$("#sectionBtn li").removeClass("active");
			$("#sectionBtn li").eq(3).addClass("active");

		}else if(scrollTopPx >= secTop[3]){
			$("html").addClass("scrolled_row3");
			$("#section3").addClass("active");
			$("#sectionBtn li").removeClass("active");
			$("#sectionBtn li").eq(2).addClass("active");

		}else if(scrollTopPx >= secTop[2]){
			$("html").addClass("scrolled_row2");
			$("#section2").addClass("active");
			$("#sectionBtn li").removeClass("active");
			$("#sectionBtn li").eq(1).addClass("active");

		}else if(scrollTopPx >= secTop[1]){
			$(".section").not("#section1").removeClass("active");

			$("html").addClass("scrolled_row1");
			$("#sectionBtn li").removeClass("active");
			$("#sectionBtn li").eq(0).addClass("active");
		}
	}
}