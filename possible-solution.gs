var names = ["name", "something", 12];
var values = [1, 2, "age"];

var obj = names.reduce(function(obj, name, i) {
  if(typeof name !== "string")
    return obj;

  obj[name] = values[i];
  return obj;
}, {});

var result = JSON.stringify(obj, null, 4);
console.log(result);




