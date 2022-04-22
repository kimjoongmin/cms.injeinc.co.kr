<!--
function info_del(sn) {
 var bool = confirm("정보를 정말 삭제하시겠습니까?");
  if(bool==false){
   return false;
  }
 return true;
}

// 약관동의 체크
function check_promis(tmp, variable)
{
	if (tmp.checked == false) {
		alert (variable + "을(를) 읽어보시고 동의하셔야 합니다.");
		tmp.focus();
		return false;
	}
	return true;
}


// 입력 체크
function check_input(tmp, variable)
{
	if (tmp.value.search(/(\S+)/) == -1) {
		alert (variable + "을(를) 입력해 주세요..");
		tmp.focus();
		tmp.select();
		return false;
	}
	return true;
}
// DOM 개체.. 
function check_in(tmp, variable)
{
	if (tmp.search(/(\S+)/) == -1) {
		alert (variable + "을(를) 입력해 주세요.");
		return false;
	}
	return true;
}
function check_input1(tmp, variable)
{
	if (tmp.value.search(/(\S+)/) == -1) {
		alert (variable + "을(를) 입력해 주세요.");
		//tmp.focus();
		//tmp.select();
		return false;
	}
	return true;
}

function check_input2(tmp, variable)
{
	if (tmp.value.search(/(\S+)/) == -1) {
		alert (variable + "를 확인해주세요.");
		//tmp.focus();
		//tmp.select();
		return false;
	}
	return true;
}

function check_input3(tmp, variable)
{
	if (tmp.value.search(/(\S+)/) == -1) {
		alert (variable + "를 선택해주세요.");
		return false;
	}
	return true;
}

// 주소 체크후 새창으로 검색창 오픈
function zip_input(tmp, variable)
{
	if (tmp.value.search(/(\S+)/) == -1) {
		alert (variable + "을(를) 입력해 주세요.");
		zip_check();
		return false;
	}
	return true;
}

// 공백 체크
function check_null(tmp, variable)
{
	if (tmp.value.indexOf(" ") != -1) {
		alert (variable + "을(를) 공백없이 입력해 주세요.");
		tmp.focus();
		tmp.select();
		return false;
	}
	return true;
}

// 한글만 입력
function isHAN(tmp, variable) 
{
    for (var i=0; i<tmp.value.length; i++){
       if(tmp.value.charCodeAt(i) <= 256){
          alert(variable + "은(는) 한글로만 입력해 주세요.");
		  tmp.focus();
		  tmp.select();
		  return false;
       }
    }
    return true;
}


// 영문만 입력
function isEng(value)
{
	for (var i=0; i<value.length; i++){
	   if(value.charCodeAt(i) <= 256){
		  return true;
	   }
	}
	return false;
}

function isEng2(tmp,variable)
{
	for (var i=0; i<tmp.value.length; i++){
	   if(tmp.value.charCodeAt(i) <= 256){
		   alert(variable + "은(는) 영어로만 입력해 주세요.");
		   return true;
	   }
	}
	return false;
}

// 숫자 체크	숫자만 사용하기
function check_number(tmp, variable) 
{
	var i;
	for (i=0; i<tmp.value.length; i++) 
	if (tmp.value.charAt(i) < '0' || tmp.value.charAt(i) > '9') {
		alert(variable + "은(는) 숫자만 입력해 주세요.");
		tmp.focus();
		tmp.select();
		return false;
	}
	return true;
}

// 실시간 숫자 체크
function NumberCheck(variable)
{
	if((event.keyCode<48)||(event.keyCode>57))
	{
		event.returnValue=false;
		alert(variable + "은(는) 숫자만 입력하세요")
	}
}

// 입력 최소길이 체크
function check_minlength(tmp, minlength, variable)
{
	if (tmp.value.length < minlength) {
		alert(variable + "이(가) 너무 짧습니다.\n다시 입력해 주세요.");
		tmp.focus();
		tmp.select();
		return false;
	}
	return true;
}



// 이미지 GIF 와 JPG(확장자) 업로드만 가능하게 하기
function img_check(tmp, variable) 
{
	var f = tmp;
	var t = tmp.value;
	var str = t.split('.');

		if((tmp.value.lastIndexOf(".jpg")==-1) && (tmp.value.lastIndexOf(".gif")==-1) && (tmp.value.lastIndexOf(".GIF")==-1) && (tmp.value.lastIndexOf(".JPG")==-1)) {
		alert(variable + "업로드할 파일은 GIF 와 JPG만 가능합니다. ")
		return false
	}
	return true
}

// 이미지 CSV 와 JPG(확장자) 업로드만 가능하게 하기
function csv_check(tmp, variable) 
{
	var f = tmp;
	var t = tmp.value;
	var str = t.split('.');

		if((tmp.value.lastIndexOf(".CSV")==-1) && (tmp.value.lastIndexOf(".csv")==-1)) {
		alert(variable + "업로드할 파일은 CSV만 가능합니다. ")
		return false
	}
	return true
}

// 이미지 pdf(확장자) 업로드만 가능하게 하기
function img_pdf_check(tmp, variable) 
{
	var f = tmp;
	var t = tmp.value;
	var str = t.split('.');

		if((tmp.value.lastIndexOf(".pdf")==-1) && (tmp.value.lastIndexOf(".PDF")==-1)) {
		alert(variable + "업로드할 파일은 PDF만 가능합니다. ")
		return false
	}
	return true
}


// 동영상 AVI / WMV (확장자) 업로드만 가능하게 하기
function img_move_check(tmp, variable) 
{
	var f = tmp;
	var t = tmp.value;
	var str = t.split('.');

		if((tmp.value.lastIndexOf(".avi")==-1) && (tmp.value.lastIndexOf(".AVI")==-1) && (tmp.value.lastIndexOf(".wmv")==-1) && (tmp.value.lastIndexOf(".WMV")==-1)) {
		alert(variable + "업로드할 파일은 AVI 와 WMV만 가능합니다. ")
		return false
	}
	return true
}




// 특수문자 제어
function CheckExceptStr(value)
{
	var exceptSTR = "!@#$%^&*()-+=\|/<>.,';}{[]~`";
	var checkSTR = value;
	var checkVA;

	for(i=0; i<checkSTR.length; i++)
	{
		var CHKstr = checkSTR.charAt(i);
		for(j=0; j<exceptSTR.length; j++)
		{
			if(CHKstr == exceptSTR.charAt(j)){
				checkVA = true;
			}
		}
	}

	if (checkVA)
	{
		alert(variable + "이(가) 너무 짧습니다.\n다시 입력해 주세요.");
	}

	if(i == checkSTR.length){
		return false;
	}
}

// 영문자와 숫자 체크
function check_stringornumber(tmp, variable)
{
	for (i = 0; i < tmp.value.length; i++) {
		if (tmp.value.charAt(i) >= '0' && tmp.value.charAt(i) <= '9')
			continue;
		else if (tmp.value.charAt(i) >= 'a' && tmp.value.charAt(i) <= 'z')
			continue;
		else if (tmp.value.charAt(i) >= 'A' && tmp.value.charAt(i) <= 'Z')
			continue;
		else {
			alert(variable + "은(는) 영문자, 숫자만 사용하실 수 있습니다.");
			tmp.focus();
			tmp.select();
			return false;
		}
	}
	return true;
}


//비밀번호 확인 체크
function check_compare(tmp1, tmp2, variable)
{	
	if (tmp1.value != tmp2.value) {
		alert(variable + "(이)가 일치하지 않습니다.\n다시 입력해 주세요.") ;
		tmp1.focus();
		tmp1.select();
		return false;
	}
	return true;
}


// 주민번호 유효성체크
function validity_resident(tmp1, tmp2, variable)
{
	if (validity_resident1(tmp1, tmp2, variable)) {
		return true;
	} else {
		if (validity_resident2(tmp1, tmp2, variable)) {
			return true;
		} else {			
			alert(variable + "은(는) 잘못된 번호입니다. 다시 입력해 주세요.");
			tmp1.value = ""
			tmp2.value = ""
			tmp1.focus();
			tmp1.select();
			return false;   
		}
	}		
}

// 선택 체크
function check_select(tmp, variable)
{
	if (tmp.options[tmp.selectedIndex].value == "") {
		alert (variable + "을(를) 선택해 주세요.");
		tmp.focus();
		return false;
	}
	return true;
}


// 이메일 체크
function check_email(tmp)
{
	if (tmp.value.search(/(\S+)@(\S+)\.(\S+)/) == -1) { 
		alert("이메일의 형식은 '계정@[호스트|도메인]'과 같이 구성되어야 합니다."); 
		tmp.focus(); 
		tmp.select(); 
		return false; 
	} 	
	for (i = 0; i < tmp.value.length; i++) {
		if (tmp.value.charAt(i) >= '0' && tmp.value.charAt(i) <= '9')
			continue;
		else if (tmp.value.charAt(i) >= 'a' && tmp.value.charAt(i) <= 'z')
			continue;
		else if (tmp.value.charAt(i) >= 'A' && tmp.value.charAt(i) <= 'Z')
			continue;
		else if (tmp.value.charAt(i) == '_' || tmp.value.charAt(i) == '-' || tmp.value.charAt(i) == '@' || tmp.value.charAt(i) == '.')
			continue;
		else {
			alert("이메일은 영문자, 숫자, 기호 '@', '.', '-' , '_' 만 사용하실 수 있습니다.");
			tmp.focus(); 
			tmp.select(); 
			return false; 
		}
	}
	if (tmp.value.charAt(0) == '_' || tmp.value.charAt(0) == '-') {
		alert("'_'와 '-'는 이메일의 첫글자로 사용하실 수 없습니다.");
		tmp.focus(); 
		tmp.select();
		return false;
	} 
	return true;
}

// 주민번호 유효성체크
function validity_resident1(tmp1, tmp2, variable)
{
	var tmp = tmp1.value + tmp2.value;
	var a = new Array(13);
	for (var i=0; i < 13; i++) {
		a[i] = parseInt(tmp.charAt(i));
	}

	if (a[2] > 1) return false;   
	if (a[4] > 3) return false;   
	if (a[6] > 4 || a[6] == 0) return false;   

	var j = a[0]*2 + a[1]*3 + a[2]*4 + a[3]*5 + a[4]*6 + a[5]*7 + a[6]*8 + a[7]*9 + a[8]*2 + a[9]*3 + a[10]*4 + a[11]*5;
	var j = j % 11;
	var k = 11 - j;

	if (k > 9) k = k % 10;

	if (k != a[12]) {
		alert("주민번호가 일치하지 않습니다.")
		return false;
	}

	return true;
}


// 외국인번호 유효성체크
function validity_resident2(tmp1, tmp2, variable)
{
	var tmp = tmp1.value + tmp2.value;
	var a = new Array(13);
	var birthday;
	var birthYear;
	var birthMonth;
	var birthDate;

	for (var i=0; i < 13; i++) {
		a[i] = parseInt(tmp.charAt(i));
	}
	
	if (a[6] == "5" || a[6] == "6") {
		birthYear = "19";
	} else if (a[6] == "7" || a[6] == "8") {
		birthYear = "20";
	} else if (a[6] == "9" || a[6] == "0") {
		birthYear = "18";
	} else {
		return false;
	}        

	birthYear += tmp.substr(0, 2);
	birthMonth = tmp.substr(2, 2) - 1;
	birthDate = tmp.substr(4, 2);
	birthday = new Date(birthYear, birthMonth, birthDate);
		
	if (birthday.getYear() % 100 != tmp.substr(0, 2) || birthday.getMonth() != birthMonth || birthday.getDate() != birthDate) return false;

	if (((a[7]*10 + a[8]) % 2) != 0) return false;

	if ((a[11] != 6) && (a[11] != 7) && (a[11] != 8) && (a[11] != 9)) return false;

	var j = a[0]*2 + a[1]*3 + a[2]*4 + a[3]*5 + a[4]*6 + a[5]*7 + a[6]*8 + a[7]*9 + a[8]*2 + a[9]*3 + a[10]*4 + a[11]*5;
	var j = j % 11;
	var k = 11 - j;

	if (k>=10) k-=10;
	k += 2;

	if (k>=10) k-=10;
	if ( k != a[12]) {
		return false;
	} else {
		return true;
	}
}


// 입력 최대길이 체크
function check_maxlength(tmp, minlength, variable)
{
	if (tmp.value.length > minlength) {
		alert(variable + "이(가) 너무 길것 같습니다.\n\n다시 입력해 주세요.");
		tmp.focus();
		tmp.select();
		return false;
	}
	return true;
}



// 입력 최소길이 체크1
function check_minlength1(tmp, minlength, variable)
{
	if (tmp.value.length < minlength) {
		alert(variable + "이상입니다.\n확인 입력해 주세요.");
		tmp.focus();
		tmp.select();
		return false;
	}
	return true;
}


function check_min_max(tmp, min_max, variable)
{
	if ((tmp.value.length < 10 ) || (tmp.value.length > 10 ))
	{
		alert(variable + "10자리를 입력해주십시오")
		tmp.focus();
		return false;
	}
	return true;
}

//-----------------------숫자를 입력하는지 확인--------------------
function checkNum(obj) {
 var str = obj.value;
 var flag = 0;

 /* 널값허용... */
 if(obj.value == "")
  return true;

 for(i=0 ; i<str.length ; i++)
  {
  if(str.charAt(i)!="0" &&
   str.charAt(i)!="1" &&
   str.charAt(i)!="2" &&
   str.charAt(i)!="3" &&
   str.charAt(i)!="4" &&
   str.charAt(i)!="5" &&
   str.charAt(i)!="6" &&
   str.charAt(i)!="7" &&
   str.charAt(i)!="8" &&
   str.charAt(i)!="9" &&
   str.charAt(i)!=" " )
   {
   alert("숫자로 정확하게 입력해 주세요... ");
   obj.value = "" ;
   obj.focus();
   return false;
   }
  }
  return true;
}//---------------------------------------------------------------------









// 숫자 체크1
function check_number1(tmp) 
{
	var i;
	for (i=0; i<tmp.value.length; i++) 
	if (tmp.value.charAt(i) < '0' || tmp.value.charAt(i) > '9') {
		alert("\n숫자만 입력가능합니다.");
		tmp.value = "";
		tmp.focus();
		tmp.select();
		return false;
	}
	return true;
}

//---------텍스트폼에서 공백 제거-------------------------
function del_blank(obj) {
  var str_in = obj.value	
	  str_out = ""
  for(i=0 ; i<str_in.length ; i++)
	if (str_in.charAt(i) != " ")	
		str_out = str_out + str_in.charAt(i)
  obj.value = str_out
}//-------------------------------------------------------





// 라디오 체크
function check_radio(tmp, variable)
{
	var checked_idx = -1;
	for (i=0; i<tmp.length; i++){
		if (tmp[i].checked)	{
			checked_idx = i;
		}
	}
	if (checked_idx < 0) {
		alert (variable + "을(를) 선택해 주세요.");
		tmp[0].focus();
		return false;
	}
	return true;
}

// 체크박스 체크
function check_checkbox(tmp, variable)
{
	var checked_cnt = 0;
	for (i=0; i<tmp.length; i++){
		if (tmp[i].checked)	{
			checked_cnt = checked_cnt + 1;
		}
	}
	if (checked_cnt <= 0) {
		alert (variable + "을(를) 체크해 주세요.");
		tmp[0].focus();
		return false;
	}
	return true;
}







// 메시지 체크
function check_message(msg)
{
	return confirm(msg);
}

// 입력값 비교 체크













// 성인 체크
function check_adult(tmp1, tmp2)
{
	var tmp = tmp1.value + tmp2.value;
	var a = new Array(13);
	for (var i=0; i < 13; i++) {
		a[i] = parseInt(tmp.charAt(i));
	}
	var birthday;
	var today;
	var birthYear;
	var birthMonth;
	var birthDate;
	if (a[6] == "1" || a[6] == "2")	{
		birthYear = "19"
	} else if (a[6] == "3" || a[6] == "4") {
		birthYear = "20";
	} else if (a[6] == "5" || a[6] == "6") {
		birthYear = "19";
	} else if (a[6] == "7" || a[6] == "8") {
		birthYear = "20";
	} else if (a[6] == "9" || a[6] == "0") {
		birthYear = "18";
	} else {
		return false;
	} 

	birthYear += tmp.substr(0, 2);
	birthMonth = tmp.substr(2, 2) - 1;
	birthDate = tmp.substr(4, 2);
	birthday = new Date(birthYear, birthMonth, birthDate);

	var today;
	var toYear;
	var toMonth;
	var toDate;

	today = new Date();
	toYear = today.getYear();
	toMonth = today.getMonth() + 1;
	toDate = today.getDate();
	today = new Date(toYear, toMonth, toDate);

	var age;
	age = (today - birthday) / (24*60*60*1000*365);

	if (age < 19) {
		alert("성인이 아닙니다.");
		return false;
	}
	return true;
}

function CKEDITOR_textarea(testareaname){
	  if(CKEDITOR.instances.content.getData() == ""){
	     alert(' 내용을 입력하세요.');
		 return false;
      } else {
         return true;
	  }
}


//if(!textarea('content')) return;
function textarea(testareaname){
 var oEditor = FCKeditorAPI.GetInstance(testareaname) ;
	 oEditortxt = oEditor.GetXHTML();
	  if(!oEditortxt){
	     alert(' 내용을 입력하세요.');
		 oEditor.Focus();
		 return false;
      } else {
         return true;
	  }
}

function SR_submit_searchFrom()
{
	var val = document.search.SearchString;


	if (!check_input(val, '검색어')) 
    {
      //alert("검색할 단어를 입력해 주세요");
      val.value="";
      val.focus();
      return false;
    }
	//document.frmSearch.submit();
	return true;
}


//가장 상단의 체크박스 선택
function menu_chooseTop(name){
	var obj = eval("document.myform."+name);
	var cnt = obj.length
	var result = false;
	for(i=0;i<cnt;i++){
		if (obj[0].checked == true){
			result = true;
		}
		obj[i].checked = result;
	}
}

//가장 상단을 재외한 체크박스 선택    
//name,시작번호,끝번호,타입(y,n,z)
//타입의 종류 y,n,z 
//y 하위체크박스가 있는 상태
//n 하위체크박스
//z 일반적인 체크박스
function menu_chooseSub(name,s,f,txt){
	var obj = eval("document.myform."+name);
		
	var result = false;
	var result2 = false;
	var result3 = false;
	var sss = parseInt(s)+1;
	var ss = parseInt(s);
	var ff = parseInt(f);
	var cnt = parseInt(obj.length)-1;

	if(txt == "y"){
		for(i=ss;i<=ff;i++){
			if (obj[ss].checked == true){
				result = true;
			}
			obj[i].checked = result;
		}
	}else if(txt == "n"){
		for(i=sss;i<=ff;i++){
			if (obj[i].checked == true){
				result2 = true;
			}
		}
		obj[s].checked = result2;
	}
	for(i=1;i<=cnt;i++){
		if (obj[i].checked == true){
			result3 = true;
		}			
	}
	obj[0].checked = result3;
}


/*********************************************************************************************************/
/* 팝업 관련 함수 ****************************************************************************************/
/*********************************************************************************************************/

/* 새창 띠우기 ==========================================================================================*/
function SR_openWindow(name, url, left, top, width, height, toolbar, menubar, statusbar, scrollbar, resizable, locations)
{

  toolbar_str = toolbar ? 'yes' : 'no';
  menubar_str = menubar ? 'yes' : 'no';
  statusbar_str = statusbar ? 'yes' : 'no';
  scrollbar_str = scrollbar ? 'yes' : 'no';
  resizable_str = resizable ? 'yes' : 'no';
  window.open(url, name, 'left='+left+',top='+top+',width='+width+',height='+height+',toolbar='+toolbar_str+',menubar='+menubar_str+',status='+statusbar_str+',scrollbars='+scrollbar_str+',resizable='+resizable_str+',location='+locations);

}

/* queyrstring 형태로 popup시키기 =======================================================================*/
// usage : onclick=""SR_goPopByQuerySTring('win_catW|pop_catW.asp?cat_no=" & rsCat_no&"|300|300|420|260|0|0|0|0|0|0')"">
	function SR_goPopByQuerySTring(aryPopparam)
	{
		var ap, newap
		var aryProperty
		var i;
		ap = aryPopparam.split("|")
		aryProperty = new Array("name","url","left","top","w","h","t","mnu","status","scroll","resize","locations")
		newap = new Array();
		for(i=0; i < aryProperty.length; i++)
		{
			newap[i] = ap[i]
		}
		
		//alert(newap[9]);
	
		newap[6] =  parseInt(newap[6]) ;
		newap[7] =  parseInt(newap[7]) ;
		newap[8] =  parseInt(newap[8]) ;
		newap[9] =  parseInt(newap[9]) ;
		newap[10] =  parseInt(newap[10]) ;
		newap[11] =  parseInt(newap[11]) ;


		SR_openWindow(newap[0], newap[1], newap[2],newap[3], newap[4], newap[5], newap[6], newap[7], newap[8], newap[9], newap[10], newap[11]);
		
	}

/* queyrstring 형태로 popup시키기 =======================================================================*/
	function SR_goPopByQuerySTring_ver01(url,param,aryPopparam, cfrm_flag, cfrm_msg)
	{
		var full_url
		var ap, newap
		var aryProperty
		var i;

		if( cfrm_flag == true )
		{
			if(!confirm(cfrm_msg) ) 
			{
				return ;
			}
		}
		
		//alert(url);
		full_url = ( param == "" ) ? url : url + "?" + param ;

		ap = aryPopparam.split("|")
		aryProperty = new Array("name","left","top","w","h","t","mnu","status","scroll","resize","locations")
		newap = new Array();
		for(i=0; i < aryProperty.length; i++)
		{
			newap[i] = ap[i]
		}
		
		//alert(newap[9]);
	
		newap[5] =  parseInt(newap[5]) ;
		newap[6] =  parseInt(newap[6]) ;
		newap[7] =  parseInt(newap[7]) ;
		newap[8] =  parseInt(newap[8]) ;
		newap[9] =  parseInt(newap[9]) ;
		newap[10] =  parseInt(newap[10]) ;

	
		SR_openWindow(newap[0], full_url, newap[1], newap[2],newap[3], newap[4], newap[5], newap[6], newap[7], newap[8], newap[9], newap[10] );
		
	}

/* form 형태로 popup시키기 =============================================================================*/
// onclick="jhj_goPopForm('/naver.com/,testForm,'300|300|420|260|0|0|0|0|0|0')"
	function SR_goPopForm(url,form,aryPopparam, cfrm_flag, cfrm_msg)
	{
		var full_url
		var ap, newap
		var aryProperty
		var i;

		if( cfrm_flag == true )
		{
			if(!confirm(cfrm_msg) ) 
			{
				return ;
			}
		}

		//full_url = ( param == "" ) ? url : url + "?" + param ;
		full_url = url;

		ap = aryPopparam.split("|")
		aryProperty = new Array("name","left","top","w","h","t","mnu","status","scroll","resize","locations")
		newap = new Array();
		for(i=0; i < aryProperty.length; i++)
		{
			newap[i] = ap[i]
		}
		
		//alert(newap[9]);
	
		newap[5] =  parseInt(newap[5]) ;
		newap[6] =  parseInt(newap[6]) ;
		newap[7] =  parseInt(newap[7]) ;
		newap[8] =  parseInt(newap[8]) ;
		newap[9] =  parseInt(newap[9]) ;
		newap[10] =  parseInt(newap[10]) ;

	
		SR_openWindow(newap[0], '', newap[1], newap[2],newap[3], newap[4], newap[5], newap[6], newap[7], newap[8], newap[9], newap[10] );
		
		form.target = newap[0];
		form.action = full_url;
		form.submit();
	}

var old='';
 
function menu(name)
{
submenu=eval("submenu_"+name+".style");
		if(old!=submenu)
	{
		if(old!='')
		{
			old.display='none';
		}
		submenu.display='block';
		old=submenu;
	}
	else
	{
		submenu.display='none';
		old='';
	}
}

//XMLHttpRequest객체를 생성해 주는 getXMLHttpRequest()함수

function getXMLHttpRequest() {
 if (window.ActiveXObject) {
  try {
   return new ActiveXObject("Msxml2.XMLHTTP");
  } catch(e) {
   try {
    return new ActiveXObject("Microsoft.XMLHTTP");
   } catch(e1) { return null; }
  }
 } else if (window.XMLHttpRequest) {
  return new XMLHttpRequest();
 } else {
  return null;
 }
}


//생성된 XMLHttpRequest객체를 저장할 전역변수
var httpRequest = null;

//XMLHttpRequest객체를 사용해서 지정한방식,지정한 URL, 첨부할 파라미터 값을 사용하여 웹 서버에 요청을 전송

function sendRequest(url, params, callback, method) {
 httpRequest = getXMLHttpRequest();
 var httpMethod = method ? method : 'GET';
 if (httpMethod != 'GET' && httpMethod != 'POST') {
  httpMethod = 'GET';
 }
 var httpParams = (params == null || params == '') ? null : params;
 var httpUrl = url;
 if (httpMethod == 'GET' && httpParams != null) {
  httpUrl = httpUrl + "?" + httpParams;
 }
 httpRequest.open(httpMethod, httpUrl, true);
 httpRequest.setRequestHeader( //컨텐트 타입 지정
  'Content-Type', 'application/x-www-form-urlencoded');
 httpRequest.onreadystatechange = callback; //readyState값이 바뀔 때 호출 될 콜백 함수 지정
 httpRequest.send(httpMethod == 'POST' ? httpParams : null); //Http요청 방식이 'POST'이면 send()함수를 통해 파라미터 파라미터 전송
}



//아이디 영어+숫자 조합체크
function pass_check(f, variable){
	var str = f.value;
		r0=/[A-z]/gi;
		r1=/[0-9]/gi;
		t0=str.search(r0);
		t1=str.search(r1);
	if(t0!=-1&&t1!=-1){
		return true;
	} else {
		alert(variable + "는(은) 영문자와 숫자 조합으로 입력하셔야 합니다.");
		f.focus();
		f.select();
		return false;
	}
}

function isHAN2(tmp, variable) 
{
    for (var i=0; i<tmp.value.length; i++){
       if(tmp.value.charCodeAt(i) >= 256){
          alert(variable + "에 한글은 사용할수 없습니다.");
		  tmp.focus();
		  tmp.select();
		  return false;
       }
    }
    return true;
}


function check_checkbox2(form, variable)
{

	var checked_cnt = 0;
	var chk_co		= document.getElementsByName("chk").length;

	for (i=0; i < chk_co; i++){
		if (chk_co==1){
			if (form.chk.checked)	{
				checked_cnt = checked_cnt + 1;
			}
		}else{
			if (form.chk[i].checked)	{
				checked_cnt = checked_cnt + 1;
			}
		}
	}
	if (checked_cnt <= 0) {
		alert (variable + "을(를) 체크해 주세요.");
		if (chk_co==1){
		form.chk.focus();
		}else{
		form.chk[0].focus();
		}
		
		return false;
	}
	return true;
}


function addBookmark(title, url){
	if(document.all) {
		//IE
		window.external.AddFavorite(url, title);
	} else if(window.chrome) {
		//Chrome
		alert('Ctrl+D 키를 누르면 즐겨찾기에 추가하실 수 있습니다.');
	} else if (window.sidebar)  {
		//firefox
		window.sidebar.addPanel(title, url, "");
	}
}

function link_go(aa){
	//top.location.href=aa;
	if(aa==''){}else{
		window.open(aa);
	}
}

//-->