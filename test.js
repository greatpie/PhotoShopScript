#include "extendscript.csv.jsx";

var DEBUG = true;
var data = CSV.toJSON('', true, ',');
for (var index in data) {
    ele = data[index]
    for (var key in ele) {
        $.writeln(ele[key]);
    }
}