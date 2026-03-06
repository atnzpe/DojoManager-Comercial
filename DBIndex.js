function indexBy(data, campo){

  const map = {};

  data.forEach(item=>{
    map[item[campo]] = item;
  });

  return map;
}