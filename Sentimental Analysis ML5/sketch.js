var table;

function preload(){
    table = loadTable('Standards Data.xls')
}

let sentiment;
let statusEl;
let submitBtn;
let inputBox;
let sentimentResult;

function setup() {
  console.log(table);
  // initialize sentiment
  sentiment = ml5.sentiment('movieReviews', modelReady);

  // setup the html environment
  statusEl = document.querySelector('#statusText')

  inputBox = document.querySelector('#inputText');
  submitBtn = document.querySelector('#submitBtn');
  sentimentResult = document.querySelector('#score');

  // predicting the sentiment on mousePressed()
  submitBtn.addEventListener('click', getSentiment);
}

setup();

function getSentiment() {
  // get the values from the input
  const text = inputBox.value;
  // make the prediction
  const prediction = sentiment.predict(text);

  // display sentiment result on html page
  sentimentResult.innerHTML = prediction.score;
}

function modelReady() {
  // model is ready
  statusEl.innerHTML = 'model loaded';
}
