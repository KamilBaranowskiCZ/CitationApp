const interval = setInterval(function() {
    wordFile = document.querySelector('input[name="word"]').files[0]
    excelFile = document.querySelector('input[name="excel"]').files[0]
    if(wordFile != undefined){
        document.getElementById('label1').innerText = wordFile.name
    }
    if(excelFile != undefined){
        document.getElementById('label2').innerText = excelFile.name
    }
    
  }, 1000);