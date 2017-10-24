function opennew(newurl,windowName,width,height)
{
var theLeft,theTop
theLeft = (screen.width-width)/2-2
theTop = (screen.height-height)/2
//window.open(newurl,windowName,'width = '+width+',height = '+height+',scrollbars = 1,left = '+theLeft+',top = '+theTop+'').focus();
window.open(newurl,windowName,'width = '+width+',height = '+height+',scrollbars = 1,status = 0,toolbar = 0,resizable = 0,left = '+theLeft+',top = '+theTop+'').focus();
//return true;
}
function opennewfull(newurl,windowName)
{
window.open(newurl,windowName,'width = '+screen.width+',height = '+(screen.height-55)+',scrollbars = 1,toolbar = 0,resizable = 0,left = 0,top = 0').focus();
}

function openModal(newUrl,windowName,width,height,controlName){
controlName.value = showModalDialog(newUrl,windowName,"dialogWidth:"+width+";dialogHeight:"+height+";center:1")
} 