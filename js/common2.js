$(function(){

    $(".section1 .animate").addClass("active");

    var swiper = new Swiper(".mainEffect", {
        effect: "cube",
        grabCursor: true,
        autoplay: {
          delay: 2500,
          disableOnInteraction: false,
        },
        pagination:{
            el:".swiper-pagination",
            clickable: true,
            renderBullet: function(index, className){
                return '<span class="' + className + '">' + (index + 1) + '</span>';
            },

        },
    });

    $(".allNav").on("click", function(e){
        e.preventDefault();
        if(!$("#gnb").is(":visible")){
            $(this).addClass("active");
            $(this).children("em").text("전체메뉴닫기")
            $("#overlay").show();
            $("#gnb").addClass("active");
        }else{
            $(this).removeClass("active");
            $(this).children("em").text("전체메뉴")
            $("#overlay").hide();
            $("#gnb").removeClass("active");
        }
    });

    $("#sectionBtn a, #gnb a").on("click",function(e){
        e.preventDefault();
        $("html").animate({scrollTop:$(this.hash).offset().top + -100},600);
        $(".allNav.active").trigger("click");
        sectionControl();
    });

});

$(window).scroll(function(){
    animate();
    headerScroll();
    sectionControl();
    mainCase();
});
$(window).load(function(){
    mainCase();
})

function mainCase(){
    let mainCaseLi = $(".mainCase div div");
    // let html = "";
    let mainCaseLiNum = mainCaseLi.length
    if($(".mainCase").hasClass("active")){
        for(let i=0;i<mainCaseLiNum;i++){
            // html = `<div><img src="/images/img_section5_${i+1}.jpg"></div>`
            // $(".mainCase>div").append(html);
            mainCaseLi.eq(i).css({"transition-delay": + (i/20) + "s","opacity":1});
        }
    }else{
        mainCaseLi.css({"transition":"all 0.01s ease","opacity":"0"}).removeAttr("style");
    }
}

function headerScroll(){
    var scrolltop = $(window).scrollTop(),
        header = $("#header");
    if(scrolltop > 10){
        header.addClass("scroll");
    }else{
        header.removeClass("scroll");
    }
}

function animate(){
    $(".animate").each(function(){
        var winScrollTop = $(window).scrollTop(),
            winHeight = $(window).height(),
            offset = $(this).offset().top;
        if(winScrollTop + winHeight > offset + 30){
            $(this).addClass("active");
        }else{
            $(this).removeClass("active");
        }
    });
}

function sectionControl(){
    $(".section").each(function(i){
        var sectionBtnLi = $("#sectionBtn li"),
            winScrollTop = $(window).scrollTop(),
            winHeight = $(window).height(),
            offset = $(this).offset().top;
        if(winScrollTop + winHeight > offset + 450){
            $(this).addClass("active");
            sectionBtnLi.removeClass("active");
            sectionBtnLi.eq(i).addClass("active");
        }else{
            $(this).removeClass("active");
        }
    });
}
