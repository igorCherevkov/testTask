// Вывод excel в консоли
let xlsx = require('xlsx');
let wb = xlsx.readFile('Sample.xlsx');
let ws = wb.Sheets['Sheet1'];

let range = xlsx.utils.decode_range(ws['!ref']);

let array = ['A','B','C','D','E','F','G','H','I','J','K','L'];

let persons = [];
for (let i = 1; i <= range.e.r + 1; i++) {
    let person = [];
    for (let j = 0; j < range.e.c + 1; j++) {
        if (ws[`${array[j]}${i}`] != undefined) {
            person[j] = ws[`${array[j]}${i}`].w;  
        } else {
            person[j] = '-------';
        }
    }
    if (i == 1) {
        persons.push(person.join('\t\t'));
    } else {
        persons.push(person.join('\t'));
    }
}

for (let i = 0; i < range.e.r; i++) {
    console.log(persons[i]);
} 
//

// Фильтрация
const readline = require('readline');
const { stdin: input, stdout: output } = require('process');

const rl = readline.createInterface({ input, output });

rl.question('По какому полю хотите фильтровать?\n1)Пол\n2)Факультет\n3)Форма обучения(очная и тп)\n4)Степень\n5)Курс\n6)Форма обучения(бюджет и тп)\n7)Гражданство\n', (answer) => {
  switch(answer) {
    case `1`: rl.question('Выберите пол: 1)мужской, 2)женский\n', (answ) => {
        if (answ == 1) {
            const filt = persons.filter(elem => elem.split('\t')[3] == 'Мужской');
            console.log(filt.join('\n'));
        } else if (answ == 2) {
            const filt = persons.filter(elem => elem.split('\t')[3] == 'Женский');
            console.log(filt.join('\n'));
        } else {
            console.log('Таких данных нет');
        }
        rl.close();
    });
        break;
    case `2`: rl.question('Выберите факультет: 1)ИКТИБ, 2)ИРТСУ, 3)ИНЭП, 4)ИУЭС\n', (answ) => {
        if (answ == 1) {
            const filt = persons.filter(elem => elem.split('\t')[5] == 'ИКТИБ');
            console.log(filt.join('\n'));
        } else if (answ == 2) {
            const filt = persons.filter(elem => elem.split('\t')[5] == 'ИРТСУ');
            console.log(filt.join('\n'));
        } else if (answ == 3) {
            const filt = persons.filter(elem => elem.split('\t')[5] == 'ИНЭП');
            console.log(filt.join('\n'));
        } else if (answ == 4) {
            const filt = persons.filter(elem => elem.split('\t')[5] == 'ИУЭС');
            console.log(filt.join('\n'));
        } else {
            console.log('Таких данных нет');
        }
        rl.close();
    });
        break;
    case `3`: rl.question('Выберите форму обучения: 1)очная, 2)заочная, 3)очно-заочная\n', (answ) => {
        if (answ == 1) {
            const filt = persons.filter(elem => elem.split('\t')[6] == 'Очная');
            console.log(filt.join('\n'));
        } else if (answ == 2) {
            const filt = persons.filter(elem => elem.split('\t')[6] == 'Заочная');
            console.log(filt.join('\n'));
        } else if (answ == 3) {
            const filt = persons.filter(elem => elem.split('\t')[6] == 'Очно-заочная');
            console.log(filt.join('\n'));
        } else {
            console.log('Таких данных нет');
        }
        rl.close();
    });
        break;
    case `4`: rl.question('Выберите степень: 1)бакалавр, 2)ДОП, 3)магистр, 4)специалист, 5)прикладной бакалавр\n', (answ) => {
        if (answ == 1) {
            const filt = persons.filter(elem => elem.split('\t')[7] == 'Бакалавр');
            console.log(filt.join('\n'));
        } else if (answ == 2) {
            const filt = persons.filter(elem => elem.split('\t')[7] == 'Дополнительная общеразвивающая программа');
            console.log(filt.join('\n'));
        } else if (answ == 3) {
            const filt = persons.filter(elem => elem.split('\t')[7] == 'Магистр');
            console.log(filt.join('\n'));
        } else if (answ == 4) {
            const filt = persons.filter(elem => elem.split('\t')[7] == 'Специалист');
            console.log(filt.join('\n'));
        } else if (answ == 5) {
            const filt = persons.filter(elem => elem.split('\t')[7] == 'Прикладной бакалавр');
            console.log(filt.join('\n'));
        } else {
            console.log('Таких данных нет');
        }
        rl.close();
    });
        break;
    case `5`: rl.question('Выберите курс: 1)1, 2)2, 3)3, 4)4, 5)5\n', (answ) => {
        if (answ == 1) {
            const filt = persons.filter(elem => elem.split('\t')[8] == '1');
            console.log(filt.join('\n'));
        } else if (answ == 2) {
            const filt = persons.filter(elem => elem.split('\t')[8] == '2');
            console.log(filt.join('\n'));
        } else if (answ == 3) {
            const filt = persons.filter(elem => elem.split('\t')[8] == '3');
            console.log(filt.join('\n'));
        } else if (answ == 4) {
            const filt = persons.filter(elem => elem.split('\t')[8] == '4');
            console.log(filt.join('\n'));
        } else if (answ == 5) {
            const filt = persons.filter(elem => elem.split('\t')[8] == '5');
            console.log(filt.join('\n'));
        } else {
            console.log('Таких данных нет');
        }
        rl.close();
    });
        break;
    case `6`: rl.question('Выберите форму обучения: 1)бюджет, 2)ПВЗ\n', (answ) => {
        if (answ == 1) {
            const filt = persons.filter(elem => elem.split('\t')[9] == 'Бюджетная основа');
            console.log(filt.join('\n'));
        } else if (answ == 2) {
            const filt = persons.filter(elem => elem.split('\t')[9] == 'Полное возмещение затрат');
            console.log(filt.join('\n'));
        } else {
            console.log('Таких данных нет');
        }
        rl.close();
    });
        break;
    case `7`: rl.question('Выберите гражданство: 1)Российская федерация, 2) Туркменистан\n', (answ) => {
        if (answ == 1) {
            const filt = persons.filter(elem => elem.split('\t')[10] == 'Российская Федерация');
            console.log(filt.join('\n'));
        } else if (answ == 2) {
            const filt = persons.filter(elem => elem.split('\t')[10] == 'Туркменистан');
            console.log(filt.join('\n'));
        } else {
            console.log('Таких данных нет');
        }
        rl.close();
    });
        break;
    }
});
//