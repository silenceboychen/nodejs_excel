//直接使用xlsx读取excel
var xlsx = require('xlsx');
var mysql = require('mysql');
var fs = require('fs');
var async = require('async');
//连接数据库
var mysql = require('mysql');
var connection = mysql.createConnection({
    host: 'localhost',
    user: 'root',
    password: '123456',
    database:'test',
    port:3306
});
connection.connect();

var book = xlsx.readFileSync('./test.xlsx'), result = {};
//拿到当前 sheet 页对象
var sheet = book.Sheets.Sheet1,
    //得到当前页内数据范围
    range = xlsx.utils.decode_range(sheet['!ref']),
    //保存数据范围数据
    row_start = range.s.r, row_end = range.e.r,
    col_start = range.s.c, col_end = range.e.c,
    rows = [], row_data, i, addr, cell;
//按行对 sheet 内的数据循环
for(;row_start<=row_end;row_start++) {
    if(row_start==0){
    	continue;
    }
    row_data = [];
    //读取当前行里面各个列的数据
    for(i=col_start;i<=col_end;i++) {
        addr = xlsx.utils.encode_col(i) + xlsx.utils.encode_row(row_start);
        cell = sheet[addr];
        //如果是链接，保存为对象，其它格式直接保存原始值
        if(cell.l) {
            row_data.push({text: cell.w, link: cell.l.Target});
        } else {
            row_data.push(cell.w);
        }
    }
    rows.push(row_data);
}
//console.log(rows);
//return;
//同步串行执行
async.eachSeries(rows,function(row,cb){
	var date=Date.parse(row[0])/1000;
	var money=row[1]*10;
	var phone=row[2];
	//console.log(date,money,phone);
	connection.query("select * from tbl_user_recharge where phone="+phone,function(err,rows,fields){
		if(err) {
			cb(err);
		}
		//console.log(rows);
		if(rows.length>0){
			money += rows[0].money;
			var id=rows[0].id;
			//console.log(money,id);
			connection.query("update tbl_user_recharge set money ="+money+",buy_time="+date+" where id="+id,function(err1,rows1){
				if(err1) {
					cb(err1);
				}
				//console.log(rows1);
				cb();
			})
		}else{
			connection.query("insert into tbl_user_recharge(phone,money,buy_time) values ("+phone+","+money+","+date+")",function(err2,rows2){
				if(err2){
					console.log(err2);
				}
				//console.log(rows2)
				cb();
			})
		}
	});
},function(err){
	console.log(err);
});
//connection.end();