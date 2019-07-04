
function f(a)
{
var x=new ActiveXObject("Scripting.FileSystemObject");
var f1=x.GetFolder(a);
var ff=new Enumerator(f1.files);
for(;!ff.atEnd();ff.moveNext()){
document.write("--"+"<a href=\""+ff.item()+"\">"+ff.item().name+"</a><br/>");
}
var fo=new Enumerator(f1.Subfolders);
for(;!fo.atEnd();fo.moveNext()){
document.write(fo.item().name+"<br/>");
sub(fo.item());
}
}
function sub(s){
var x=new ActiveXObject("Scripting.FileSystemObject");
var f1=x.GetFolder(s);
var ff=new Enumerator(f1.files);
for(;!ff.atEnd();ff.moveNext()){
document.write("--"+"<a href=\""+ff.item()+"\">"+ff.item().name+"</a><br/>");
}
var fo=new Enumerator(f1.Subfolders);
for(;!fo.atEnd();fo.moveNext()){
document.write(fo.item().name+"<br/>");
sub(fo.item());
}
}
