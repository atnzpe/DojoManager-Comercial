const CACHE_TIME = 300;

function getTable(nomeAba){

  const cache = CacheService.getScriptCache();
  const key = "TABLE_" + nomeAba;

  const cached = cache.get(key);

  if(cached){
    return JSON.parse(cached);
  }

  const data = lerTabelaDinamica(nomeAba);

  cache.put(key, JSON.stringify(data), CACHE_TIME);

  return data;
}