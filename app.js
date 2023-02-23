const express = require ('express');
const { json } = require('body-parser');
const { urlencoded } = require('express');
const generateDeck = require('./routes/generateDeck');
const removeSlides = require('./routes/removeSlides');
const replaceTextSolvemate = require('./routes/replaceTextSolvemate');
const replaceTextCoffee1 = require('./routes/replaceTextCoffee1');
const replaceTextCoffee2 = require('./routes/replaceTextCoffee2');
const replaceTextCoffee3 = require('./routes/replaceTextCoffee3');
const replaceTextCoffee4 = require('./routes/replaceTextCoffee4');
const replaceTextCoffee5 = require('./routes/replaceTextCoffee5');
const morgan = require('morgan');

const app = express();

app.use(json());
app.use(urlencoded({extended:true}));
app.use(morgan('tiny'))

app.use(generateDeck);
app.use(removeSlides);
app.use(replaceTextSolvemate);
app.use(replaceTextCoffee1);
app.use(replaceTextCoffee2);
app.use(replaceTextCoffee3);
app.use(replaceTextCoffee4);
app.use(replaceTextCoffee5);

module.exports = app;
