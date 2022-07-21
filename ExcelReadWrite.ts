const xlsx = require('xlsx');

    function ReadData(path) {
        try {
            let file = xlsx.readFile(path)
            let sheets = file.SheetNames
            let data = []
            
            for (let i = 0; i < sheets.length; i++) {
                let temp = xlsx.utils.sheet_to_json(file.Sheets[file.SheetNames[i]])

                for(let i=0; i < temp.length; i++) {
                    data.push(temp[i])
                }
            }
           
            return data
        } catch(ex) {
            console.log('Error occured in ReadData', ex)
        }
    }

    function WriteData(readResult) {
      let writeResults = readResult

        // Conversion Logic 
        if(readResult.length > 0) {
            for(let i=0; i < readResult.length; i++) {
                if (!readResult[i]['Required Tasks'].includes('&', '|', '(', ')',',')) {
                    let split = [readResult[i]['Required Tasks']]
                    split.push('dummy')
                    let objOR = { "missing_some": [1, split] }
                    writeResults[i]['Rules'] = JSON.stringify(objOR)
                } 

             let columnSplit = readResult[i]['Required Tasks'].split('')
             let found = false
          
               for(let j=0; j < columnSplit.length; j++) {
                let obj ={}
                 if((columnSplit[j] === '&' || columnSplit[j] === ',') && !readResult[i]['Required Tasks'].includes('(',')') && !found) {
                    if(!readResult[i]['Required Tasks'].includes('|')) {
                        let split = readResult[i]['Required Tasks'].split('&')
                        obj = {"missing": split }
                        writeResults[i]['Rules'] = JSON.stringify(obj)
                        found = true
                    } else {
                        
                        let andIndex = readResult[i]['Required Tasks'].indexOf('&')
                        let OrIndex = readResult[i]['Required Tasks'].indexOf('|')
                        if(OrIndex < andIndex) {
                            let split = readResult[i]['Required Tasks'].split('&')
                            let objwithORand ={"if":[
                                {"missing_some":[1,[]]},
                                {"missing":[[]]},
                                "OK"
                              ]
                            }
    
                            split.forEach(element => {
                                if(element.includes('|')) {
                                   let d = element.split('|')
                                   d.push('dummy')
                                  objwithORand.if[0]['missing_some'][1] = d
                                  
                                } else {
                                    objwithORand.if[1]['missing'] = [element]
                                }
                            });
    
                            writeResults[i]['Rules'] = JSON.stringify(objwithORand)
                            found = true
                        } else {
                            let split = readResult[i]['Required Tasks'].split('|')
                            let objwithORand ={"if":[
                                {"missing":[[]]},
                                {"missing_some":[1,[]]},
                                "OK"
                              ]
                            }
    
                            split.forEach(element => {
                                if(element.includes('&')) {
                                   let d = element.split('&')
                                  objwithORand.if[0]['missing'] = d
                                  
                                } else {
                                    objwithORand.if[1]['missing_some'][1] = [element, 'dummy']
                                }
                            });
    
                            writeResults[i]['Rules'] = JSON.stringify(objwithORand)
                            found = true
                        }
                    }
                 } else if (columnSplit[j] === '|' && !readResult[i]['Required Tasks'].includes('(',')') && !found) {
                    if(!readResult[i]['Required Tasks'].includes('&')) {
                         let split = readResult[i]['Required Tasks'].split('|')
                         split.push('dummy')
                         let objOR = {"missing_some":[1,split]}
                         writeResults[i]['Rules'] = JSON.stringify(objOR)
                         found = true
                    } else {
                        let andIndex = readResult[i]['Required Tasks'].indexOf('&')
                        let OrIndex = readResult[i]['Required Tasks'].indexOf('|')
                        if(OrIndex < andIndex) {
                            let split = readResult[i]['Required Tasks'].split('|')
                            let objwithORand ={"if":[
                                {"missing_some":[1,[]]},
                                {"missing":[[]]},
                                "OK"
                              ]
                            }
    
                            split.forEach(element => {
                                if(element.includes('&')) {
                                   let d = [split[0], element.split('&')[0]]//element.split('&')
                                   d.push('dummy')
                                  objwithORand.if[0]['missing_some'][1] = d
                                  objwithORand.if[1]['missing'] = [element.split('&')[1]]
                                }
                            });
    
                            writeResults[i]['Rules'] = JSON.stringify(objwithORand)
                            found = true
                        } else {
                            let split = readResult[i]['Required Tasks'].split('|')
                            let objwithORand ={"if":[
                                {"missing":[[]]},
                                {"missing_some":[1,[]]},
                                "OK"
                              ]
                            }
    
                            split.forEach(element => {
                                if(element.includes('&')) {
                                   let d = element.split('&')
                                  objwithORand.if[0]['missing'] = d
                                  
                                } else {
                                    objwithORand.if[1]['missing_some'][1] = [element, 'dummy']
                                }
                            });
    
                            writeResults[i]['Rules'] = JSON.stringify(objwithORand)
                            found = true
                        }
                   
                    }
                    
                 } else if ((columnSplit[j] === '(' || columnSplit[j] === ')') && !found) {
                    let andIndex = readResult[i]['Required Tasks'].indexOf('&')
                    let OrIndex = readResult[i]['Required Tasks'].indexOf('|')
                    if(OrIndex < andIndex) {
                        
                        let split = readResult[i]['Required Tasks'].split('|')
                        let objwithORand ={"if":[
                            {"missing_some":[1,[]]},
                            {"missing":[[]]},
                            "OK"
                          ]
                        }

                        split.forEach(element => {
                            if(element.includes('&')) {
                              objwithORand.if[0]['missing_some'][1] = [split[0], 'dummy']
                              let val = split[1].replace(/[()]/g, '')
                            
                              objwithORand.if[1]['missing'] = val.split('&')
                            }
                        });

                        writeResults[i]['Rules'] = JSON.stringify(objwithORand)
                        found = true

                    }
                 }
               }
            }
        }

       console.log('writeResults',writeResults)
       
       // Write converted logic into excel output
       var fs = require('fs');

       var writeStream = fs.createWriteStream("EWNworkstreamAutomationOutput.xls");
        
       var header="Inspection"+"\t"+"Required Tasks"+"\t"+"Rules"+"\n";
       writeStream.write(header)
        var data = ''
        for (let i = 0; i < writeResults.length; i++) {
            data = data + writeResults[i].Inspection + '\t' + writeResults[i]['Required Tasks'] + '\t' + writeResults[i].Rules + '\n';
        }

        writeStream.write(data);
        writeStream.close();
    }

    var path = './EWNworkstreamAutomationInput.xlsx'
    var readResult = ReadData(path)
    var writeResult = WriteData(readResult)

