// const lez = require('./output/index');
// console.log(lez.sayName({
//     animalName : 'tro',
//     animalBread: 'pug',
//     birdName: 'bruno',
//     birdBread : 'parrot'
// }));

const excel = require('./output/excel');

let exworkbook = new excel.ExcelWorkbook();

let exsheet = new excel.ExcelWorksheet(exworkbook.getWorkbook(), 'NewSheet',{
    properties :{
        tabColor: {
            argb: 'FFFF0000'
        },
        showGridLines : false
    }
});

exsheet.addHeaders([
    {header : 'S.NO.', key : 'sno', width: 20},
    {header : 'Items', key : 'item', width: 30},
    {header : 'Quantity', key : 'qty', width: 30},
    {header : 'Price of Product', key : 'price', width: 40}
])

exsheet.colorHeader({
    argb : 'FFFF0000'
})



exsheet.addData([
    {
        "sno" : "1",
        "item" : "Cold Drinks",
        "qty" : "5",
        "price" : "50"
    },
    
    {
        "sno" : "2",
        "item" : "Hand Wash",
        "qty" : "1",
        "price" : "150"
    },
    
    {
        "sno" : "3",
        "item" : "Soaps",
        "qty" : "10",
        "price" : "100"
    },
    
    {
        "sno" : "4",
        "item" : "Dish Wash",
        "qty" : "1",
        "price" : "300"
    }
    
])

exsheet.border({
    top : {
        style : 'thin'
    },
    bottom : {
        style : 'thin'
    },
    left : {
        style : 'thin'
    },
    right : {
        style : 'thin'
    }
})

if(exworkbook.saveWorkbook('./demoFile.xlsx')){
    console.log('file has been saved ');
}
else{
    console.log('file cannot be saved ...');
}
