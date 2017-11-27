var exec = require('child_process').exec,
    fs = require('fs');
    path = require('path');


var ppt2png = function(input, output, callback) {
  if(path.extname(input) == ".ppt" || path.extname(input) == ".pptx"){
    var pdfPath = output;
    exec('unoconv -f pdf -o ' + pdfPath + '.pdf ' + input, 
    function( error, stdout, stderr) {
      if (error) {
        callback(error);
        return;
      }
      pdf2png(pdfPath+'.pdf', output, function(err){
        fs.unlink(pdfPath+'.pdf', function(err) {
          if(err) {
            console.log(err);
          }

        });

        callback(err);       
      });
    });
  }
  else if(path.extname(input) == ".pdf"){
    var pdfPath = input.substr(0, input.lastIndexOf('.')) || input;

    pdf2png(pdfPath+'.pdf', output, function(err){
      fs.unlink(pdfPath+'.pdf', function(err) {
        if(err) {
          console.log(err);
        }

      });

      callback(err);       
    });

  }  
}

var pdf2png = function(input, output, callback) {
  exec('convert -resize 1200x675 -background white -gravity center -extent 1200x675 -density 200 ' + input + ' ' + output+'.png', 
    function (error, stdout, stderr) {
      if (error) {
        callback(error);
      } else {
        callback(null);
      }
    });
}

// ppt to jpg by unoconv directly
var ppt2jpg = function(input, output) {
  exec('unoconv -f html -o ' + output + '/ ' + input, 
      function( error, stdout, stderr) {
        console.log('unoconv stdout: ', stdout);
        console.log('unoconv stderr: ', stderr);
        if (error !== null) {
          console.log('unoconv err: ', error);
        } else {
          exec('rm ' + output + '/*.html ' + output + '/' + output, 
            function(){
              if(error !== null) {
                console.log('rm err: ', error);
              }
            });
        }
      });
}

module.exports = ppt2png;
