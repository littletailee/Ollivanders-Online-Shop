

//
// ����¼����д�Ƿ�Ϸ�
//
function CheckLoginForm()
{ 
	if (document.UserLogin.UserName.value  == "")
	{
		alert("input ID");
		document.UserLogin.UserName.focus();
		return false;
	}
	if (document.UserLogin.Password.value  == "") 
	{
	alert("input PSW��");
	document.UserLogin.Password.focus();
	return false; 
	}

	UserLogin.submit();
	return true;
}

//��������fucCheckNUM
//���ܽ��ܣ�����Ƿ�Ϊ����
//����˵����Ҫ��������
//����ֵ��1Ϊ�����֣�0Ϊ��������
function fucCheckNUM(NUM)
{
	var i,j,strTemp;
	strTemp = "0123456789";
	if ( NUM.length ==  0)
		return 0
	for (i = 0;i<NUM.length;i++)
	{
		j = strTemp.indexOf(NUM.charAt(i));	
		if (j == -1)
		{
		//˵�����ַ���������
			return 0;
		}
	}
	//˵��������
	return 1;
}

function checknum(theInput) 
{
 	if ((fucCheckNUM(theInput.value) == 0) )
	{	
	   theInput.value = "";
		//theform.newprice.focus();
		return false;
	}
}


function jumpTo(i){
if(i == 1){
	this.document.location = "<%=thisUrl%>";}
if(i == 2){
	this.document.location = "<%=thisUrl%>&page=<%=page-1%>";}
if(i == 3){
	this.document.location = "<%=thisUrl%>&page=<%=page+1%>";}
if(i == 4){
	this.document.location = "<%=thisUrl%>&page=<%=rsObj.pageCount%>";}
}
//-->


function clean(){ 
	if (confirm("clear cart?") == 1){
	window.location.href = "shopCart.asp?clear=yes"}
}

function checkNumNull(theform) {
	if (theform.value == "") {	
	   alert("input number");
		//theform.newprice.focus();
		theform.focus();
		return false;
	}
}
//-->

//��������chksafe
//���ܽ��ܣ�����Ƿ���"'",'\\',"/"
//����˵����Ҫ�����ַ���
//����ֵ��0����  1������
function chksafe(a)
{	
	return 1;
/*	fibdn  =  new Array ("'" ,"\\", "��", ",", ";", "/");
	i = fibdn.length;
	j = a.length;
	for (ii = 0;ii<i;ii++)
	{	for (jj = 0;jj<j;jj++)
		{	temp1 = a.charAt(jj);
			temp2 = fibdn[ii];
			if (tem';p1 == temp2)
			{	return 0; }
		}
	}
	return 1;
*/	
}

//��������chkemail
//���ܽ��ܣ�����Ƿ�ΪEmail Address
//����˵����Ҫ�����ַ���
//����ֵ��0������  1����
function chkemail(a)
{	var i = a.length;
	var temp  =  a.indexOf('@');
	var tempd  =  a.indexOf('.');
	if (temp > 1) {
		if ((i-temp) > 3){
			
				if ((i-tempd)>0){
					return 1;
				}
			
		}
	}
	return 0;
}

//��������fucPWDchk
//���ܽ��ܣ�����Ƿ��з����ֻ���ĸ
//����˵����Ҫ�����ַ���
//����ֵ��0������ 1��ȫ��Ϊ���ֻ���ĸ
function fucPWDchk(str)
{
  var strSource  = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ";
  var ch;
  var i;
  var temp;
  
  for (i = 0;i<= (str.length-1);i++)
  {
  
    ch  =  str.charAt(i);
    temp  =  strSource.indexOf(ch);
    if (temp == -1) 
    {
     return 0;
    }
  }
  if (strSource.indexOf(ch) == -1)
  {
    return 0;
  }
  else
  {
    return 1;
  } 
}

//��������fucCheckNUM
//���ܽ��ܣ�����Ƿ�Ϊ����
//����˵����Ҫ��������
//����ֵ��1Ϊ�����֣�0Ϊ��������
function fucCheckNUM(NUM)
{
	var i,j,strTemp;
	strTemp = "0123456789";
	if ( NUM.length ==  0)
		return 0
	for (i = 0;i<NUM.length;i++)
	{
		j = strTemp.indexOf(NUM.charAt(i));	
		if (j == -1)
		{
		//˵�����ַ���������
			return 0;
		}
	}
	//˵��������
	return 1;
}

//��������chkspc
//���ܽ��ܣ�����Ƿ��пո�
//����˵����Ҫ�����ַ���
//����ֵ��0����  1������
function chkspc(a)
{
	var i = a.length;
	var j  =  0;
	var k  =  0;
	while (k<i)
	{
		if (a.charAt(k) !=  " ")
			j  =  j+1;
		k  =  k+1;
	}
	if (j == 0)
	{
		return 0;
	}
	
	if (i!= j)
	{ return 2; }
	else
	{
		return 1;
	}
}

//��������fucCheckTEL
//���ܽ��ܣ�����Ƿ�Ϊ�绰����
//����˵����Ҫ�����ַ���
//����ֵ��1Ϊ�ǺϷ���0Ϊ���Ϸ�
function fucCheckTEL(TEL)
{
	var i,j,strTemp;
	strTemp = "0123456789-()# ";
	for (i = 0;i<TEL.length;i++)
	{
		j = strTemp.indexOf(TEL.charAt(i));	
		if (j == -1)
		{
		//˵�����ַ����Ϸ�
			return 0;
		}
	}
	//˵���Ϸ�
	return 1;
}

//��������fucCheckLength
//���ܽ��ܣ�����ַ����ĳ���
//����˵����Ҫ�����ַ���
//����ֵ������ֵ
function fucCheckLength(strTemp)
{
	var i,sum;
	sum = 0;
	for(i = 0;i<strTemp.length;i++)
	{
		if ((strTemp.charCodeAt(i)>= 0) && (strTemp.charCodeAt(i)<= 255))
			sum = sum+1;
		else
			sum = sum+2;
	}
	return sum;
}

function form1_onsubmit() 
{
	if (chkspc(document.form1.name.value) == 0)
	{	alert("ipnut name");
		document.form1.name.focus();
		return false;
	}
	
	if ((window.form1.sex[0].checked  ==  0)  &&  (window.form1.sex[1].checked  ==  0  ))
	{	alert("choose sex");
		return false;
	}
	
	if ((chksafe(document.form1.name.value) == 0)||(fucCheckLength(document.form1.name.value)>20))
	{	alert("name wrong");
		document.form1.name.focus();
		return false;		
	}
	
	if (fucCheckLength(document.form1.pwd.value)<4) 
	{	alert("at least 4 characters in PSW")
		document.form1.pwd.focus();
		return false;
	}
	
	if ((chksafe(document.form1.pwd.value) == 0)||(fucCheckLength(document.form1.pwd.value)>18))
	{	alert("wrong PSW form")
		document.form1.pwd.focus();
		return false;		
	}
	
	
	if (document.form1.PasswordConfirm.value!= document.form1.pwd.value)
	{
		alert ("confirm PSW");
		document.form1.PasswordConfirm.value = '';    
		document.form1.pwd.value = '';
		document.form1.pwd.focus();
		return false;
	}
	

	if ((chkspc(document.form1.email.value) == 0) || (chkemail(document.form1.email.value) == 0))
	{	alert ("wrong email");
		document.form1.email.focus();
		return false;
	}
	if ((chksafe(document.form1.email.value) == 0)||(fucCheckLength(document.form1.email.value)>40))
	{	alert ("input email");
		document.form1.email.focus();
		return false;		
	}

	if ((document.form1.phone.value == '') || (chkspc(document.form1.phone.value) == 0) || (fucCheckLength(document.form1.phone.value)>30)||(fucCheckTEL(document.form1.phone.value) == 0))
	{
		alert("wrong phone");
		document.form1.phone.focus();
		return false;
	}
	
	if (chkspc(document.form1.address.value) == 0)
	{	alert ("input address");
		document.form1.address.focus();
		return false;
	}
	if ((chksafe(document.form1.address.value) == 0)||(fucCheckLength(document.form1.address.value)>200))
	{	alert ("input address");
		document.form1.address.focus();
		return false;		
	}
	if (chkspc(document.form1.zipcode.value) == 0)
	{	alert ("input zip code");
		document.form1.zipcode.focus();
		return false;
	}
	if ((chksafe(document.form1.zipcode.value) == 0)||(fucCheckLength(document.form1.zipcode.value)>15))
	{	alert ("input zip code");
		document.form1.zipcode.focus();
		return false;		
	}
	
	for (lgth = 0;lgth<= document.form1.pwd.value.length;lgth++)
		{	if ( (document.form1.pwd.value.charCodeAt(lgth)>128)  || (document.form1.pwd.value.charAt(lgth) == "'") )
			{	alert("PSW is English only");
				document.form1.pwd.focus();
				return false;
			}
		}
}

function checkMe(theForm){
	if(theForm.name.value == ""){
			alert("name can't be empty");
			return false;
	}
	return true;
}