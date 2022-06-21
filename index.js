//import pptxgen from "pptxgenjs";
const pptxgen = require('pptxgenjs');
var fs = require('fs');

const outDirName = 'out';
if (!fs.existsSync(outDirName)) {
    fs.mkdirSync(outDirName);
}
function getDataForYearMonth() {

      const yearMonthDirectoryPath = 'data_files/2022_01/';
     
      const statisticsJsonStr = fs.readFileSync(yearMonthDirectoryPath + 'statistics.json');
      const successfulAuthorizationsJsonStr = fs.readFileSync(yearMonthDirectoryPath + 'successful_authorizations.json');
      const usersJsonStr = fs.readFileSync(yearMonthDirectoryPath + 'users.json');
     
      const statistics = JSON.parse(statisticsJsonStr);
      const successfulAuthorizations = JSON.parse(successfulAuthorizationsJsonStr);
      const users = JSON.parse(usersJsonStr);

const myDateStr = myDate.ToLoacaleString('ru-ru', { month: 'long', calendar: islamic-civil });
const myDate = new Date();
myDate.SetMonth(6);

    return {
        statistics: statistics,
        successfulAuthorizations: successfulAuthorizations,
        users: users
    }
}
let pres = new pptxgen();

let slide = pres.addSlide('Period');

slide.addImage({ path: "C:/Users/dfomin/Documents/Приложение Статистика/Generator-master/Presentation_settings/TGC-1/img1.png",  y: 0.1, w: 9.9, h: 0.61 });
slide.addImage({ path: "C:/Users/dfomin/Documents/Приложение Статистика/Generator-master/Presentation_settings/TGC-1/img2.png", y: 0.7, w: 10, h: 3.3 });
slide.addText("Аналитический отчет по работе модуля оптимизации", { 
   w: '100%',
   x: '5%',
   y: '85%',
    color: "2683c6",
	fontface: 'Museo Sans Cyrl 100',
	fontsize: '28',
    fill: { color: "F1F1F1" },
    align: pres.AlignH.center,
	bold: 'true',
	});
slide.addText("Отчетный период:  ", { 
   w: '100%',
   x: '-10%',
   y: '92%',
    color: "2683c6",
	fontface: 'Museo Sans Cyrl 100',
	fontsize: '18',
	align: pres.AlignH.center,

});
slide.addText( "new Date(month)", { 
   w: '100%',
   x: '10%',
   y: '92%',
    color: "2683c6",
	fontface: 'Museo Sans Cyrl 100',
	fontsize: '18',
align: pres.AlignH.center 
});


let slide2 = pres.addSlide('Soderzhanie');
title: "Содержание",
background = { color: "FFFFFF" },
objects = [
        { line: { x: 3.5, y: 1.0, w: 6.0, line: { color: "0088CC", width: 5 } } },
        { rect: { x: 0.0, y: 5.3, w: "100%", h: 0.75, fill: { color: "F1F1F1" } } },
        { text: { text: "Status Report", options: { x: 3.0, y: 5.3, w: 5.5, h: 0.75 } } },
        { image: { x: 11.3, y: 6.4, w: 1.67, h: 0.75, path: "C:/Users/dfomin/Documents/Приложение Статистика/Generator-master/Presentation_settings/TGC-1/img1.png" } },
    ],
    slideNumber  = { x: 0.3, y: "90%" };



slide2.addImage({ path: "C:/Users/dfomin/Documents/Приложение Статистика/Generator-master/Presentation_settings/TGC-1/img1.png",  y: 0.1, w: 9.9, h: 0.61 });
slide2.addText("Содержание отчета", {
   w: '50%',
    y: '50%',
	fontface: 'Museo Sans Cyrl 100',
	fontsize: "28",
    color: "2683c6",
    fill: { color: "F1F1F1" },
    align: pres.AlignH.center,
	});

let slide3 = pres.addSlide('AvtorizaciaRaschety');

slide3.addImage({ path: "C:/Users/dfomin/Documents/Приложение Статистика/Generator-master/Presentation_settings/TGC-1/img1.png",  y: 0.1, w: 9.9, h: 0.61 });
slide3.addText("Статистика использования модуля оптимизации", {
    x: 1.5,
    y: 1.5,
    color: "2683c6",
    fill: { color: "F1F1F1" },
    align: pres.AlignH.center,
});





let slide4 = pres.addSlide('Greenmine');

slide4.addImage({ path: "C:/Users/dfomin/Documents/Приложение Статистика/Generator-master/Presentation_settings/TGC-1/img1.png",  y: 0.1, w: 9.9, h: 0.61 });
slide4.addText("Статистика запросов в Техническую поддержку по работе продукта", {
    x: 1.5,
    y: 1.5,
    color: "2683c6",
    fill: { color: "F1F1F1" },
    align: pres.AlignH.center,
});

let slide5 = pres.addSlide('DopAnalitica');

slide5.addImage({ path: "C:/Users/dfomin/Documents/Приложение Статистика/Generator-master/Presentation_settings/TGC-1/img1.png",  y: 0.1, w: 9.9, h: 0.61 });
slide5.addText("Дополнительная аналитика вне Сопровождения", {
    x: 1.5,
    y: 1.5,
    color: "2683c6",
    fill: { color: "F1F1F1" },
    align: pres.AlignH.center,
});

let slide6 = pres.addSlide('DopFuncional');

slide6.addImage({ path: "C:/Users/dfomin/Documents/Приложение Статистика/Generator-master/Presentation_settings/TGC-1/img1.png",  y: 0.1, w: 9.9, h: 0.61 });
slide6.addText("Дополнительный функционал, реализованный в Отчетный период", {
    x: 1.5,
    y: 1.5,
    color: "2683c6",
    fill: { color: "F1F1F1" },
    align: pres.AlignH.center,
});

let slide7 = pres.addSlide('Problema');

slide7.addImage({ path: "C:/Users/dfomin/Documents/Приложение Статистика/Generator-master/Presentation_settings/TGC-1/img1.png",  y: 0.1, w: 9.9, h: 0.61 });
slide7.addText("Перечень проблем, вопросов, пожеланий", {
    x: 1.5,
    y: 1.5,
    color: "2683c6",
    fill: { color: "F1F1F1" },
    align: pres.AlignH.center,
	
	
});



let slide8 = pres.addSlide('KonecSlayda');

slide8.addImage({ path: "C:/Users/dfomin/Documents/Приложение Статистика/Generator-master/Presentation_settings/TGC-1/img1.png",  y: 0.1, w: 9.9, h: 0.61 });
slide8.addImage({ path: "C:/Users/dfomin/Documents/Приложение Статистика/Generator-master/Presentation_settings/TGC-1/8/img-1-str-8.png",  y: 2, w: 5.2, h: 3.5 });

slide8.addText("Надеемся, что выполненная работа повысит качество и удобство пользования продуктом нашей компании. ", {
   w: '45%',
   h: '30%',
   x: '50%',
   y: '45%',
    color: "2683c6",
	fontface: 'Museo Sans Cyrl 100',
	fontsize: '28',
	align: pres.AlignH.center,
});
slide8.addText("Будем благодарны за обратную связь!", {
   w: '60%',
   h: '30%',   
   x: '42%',
   y: '65%',
    color: "2683c6",
	fontface: 'Museo Sans Cyrl 100',
	fontsize: '28',
	align: pres.AlignH.center,
});

console.log('Presentacia created.');
pres.writeFile({ fileName: "out/Presentation-TGC-1.pptx" }); 