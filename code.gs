function main() {
  Logger.log(convertJsonFromHash(hash))
  Logger.log(convertJsonFromHash(hashInArray))
}

let hash = {
  key1: 'value1',
}

let hashInArray = [{
  key2: 'value2',
}]

function convertJsonFromHash(hash) {
  return JSON.stringify(hash)
}
