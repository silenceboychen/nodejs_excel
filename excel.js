//使用node-xlsx读取excel,实际上是对excel的又一层封装
var xlsx = require('node-xlsx');
var fs = require('fs');
//读取文件内容
var obj = xlsx.parse(__dirname+'/test.xlsx');
// console.log(obj);
var excelObj=obj[0].data;
console.log(excelObj);

var data = [];
for(var i in excelObj){
    var arr=[];
    var value=excelObj[i];
    for(var j in value){
        arr.push(value[j]);
    }
    data.push(arr);
}
console.log(data);
var buffer = xlsx.build([
 {
     name:'sheet1',
     data:data
 }
]);

//将文件内容访问新的文件中
fs.writeFileSync('test1.xlsx',buffer,{'flag':'w'});


//var data1 = [
//['name', 'age']
//];
//
////写入excel之后是一个三行两列的表格
//var data2 = [
//    ['name', 'age'], 
//    ['zhang san', '10'], 
//    ['li si', '11']
//];
//
//var buffer = xlsx.build([
//    {
//        name:'sheet1',
//        data:data1
//    }, {
//        name:'sheet2',
//        data:data2
//    }
//]);

//fs.writeFileSync('book.xlsx', buffer, {'flag':'w'}); // 如果文件存在，覆盖
