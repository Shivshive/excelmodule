// const lez = require('./output/index');
// console.log(lez.sayName({
//     animalName : 'tro',
//     animalBread: 'pug',
//     birdName: 'bruno',
//     birdBread : 'parrot'
// }));

const excel = require('./output/excel');

let exworkbook = new excel.exWorkbook();

let exsheet = new excel.exWorksheet(exworkbook.getWorkbook(), 'NewSheet',{
    properties :{
        tabColor: {
            argb: 'FFFF0000'
        },
        showGridLines : false
    },
    views : [
        {
            showRuler : true,
            showGridLines : false
        }
    ]
});

exsheet.addHeaders([
    {header : 'Name', key : 'name', width: 20},
    {header : 'Sal-18', key : 's18', width: 10},
    {header : 'Sal-19', key : 's19', width: 10},
    {header : 'Sal-20', key : 's20', width: 10},
    {header : 'Total_Sal_Count', key: 'Total_Sal_Count', width : 20}
])

exsheet.colorHeader({
    argb : 'FFFF0000'
})

let data = [
    {
        "name" : "Bobby Singer",
        "s18" : 12,
        "s19" : 30,
        "s20" : 30
    },
    {
        "name" : "John Berry Alan",
        "s18" : 10,
        "s19" : 20,
        "s20" : 10
    },
    {
        "name" : "Melissa Adhock",
        "s18" : 20,
        "s19" : 40,
        "s20" : 10
    },
    {
        "name" : "Jim Karry",
        "s18" : 20,
        "s19" : 10,
        "s20" : 15
    },
        
    ]


exsheet.addDataWithRowTotal(data,'Total_Sal_Count','Sal-18','Sal-19','Sal-20');

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
