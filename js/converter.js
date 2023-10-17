let selectedFile

document.getElementById('input').addEventListener('change', (event) => {
     selectedFile = event.target.files[0]
})



document.getElementById('button').addEventListener('click', () => {
     if (selectedFile) {
          let fileReader = new FileReader()
          fileReader.readAsBinaryString(selectedFile)
          fileReader.onload = (event) => {
               let data = event.target.result
               let workbook = XLSX.read(data, { type: 'binary' })
               // console.log(workbook)
               workbook.SheetNames.forEach(sheet => {
                    let rowObject = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheet])


                    //===================== STATUS DA ORDEM =============

                    const statusEmRota = item => item.Status === 'em rota'
                    const statusConcluida = item => item.Status === 'Concluída'
                    const statusIniciada = item => item.Status === 'Iniciada'
                    const statusNaoIniciada = item => item.Status === 'Não Iniciada'
                    const statusNaoConcluida = item => item.Status === 'Não Concluída'

                    //=======================FILTROS 



                    const conIniNin = item => {
                         if (item.Status === 'Concluída' || item.Status === 'Iniciada' || item.Status === 'Não Iniciada')
                              return item
                    }

                    const metalico = (item) => {
                         if (item['Habilidades de Trabalho'] === 'Reparo Linha(1/100)' ||
                              item['Habilidades de Trabalho'] === 'Reparo Banda(1/100)' ||
                              item['Habilidades de Trabalho'] === 'Reparo TV(1/100)')

                              return item
                    }

                    const gpon = (item) => {
                         if (item['Habilidades de Trabalho'] === 'Reparo Banda FB Alto Valor(1/100)' ||
                              item['Habilidades de Trabalho'] === 'Reparo Banda FB(1/100)' ||
                              item['Habilidades de Trabalho'] === 'Reparo Linha FB Alto Valor(1/100)' ||
                              item['Habilidades de Trabalho'] === 'Reparo Linha FB(1/100)' ||
                              item['Habilidades de Trabalho'] === 'Reparo TV FB(1/100)' ||
                              item['Habilidades de Trabalho'] === 'Reparo TV FB Alto Valor(1/100)' ||
                              item['Habilidades de Trabalho'] === 'Reparo FTTA(1/100)' ||
                              item['Habilidades de Trabalho'] === '' &&
                              (item["Tipo de Atividade"] === 'Defeito Banda Larga' ||
                                   item["Tipo de Atividade"] === 'Defeito Banda Larga' ||
                                   item["Tipo de Atividade"] === 'Defeito Banda Larga'
                              )
                         )


                              return item
                    }


                    const prev = (item => {
                         if ((item['Detalhe da Atividade'].charAt(0) === 'P' || item['Detalhe da Atividade'].charAt(0) === 'p') &&
                              (item['Detalhe da Atividade'].charAt(1) === 'R' || item['Detalhe da Atividade'].charAt(1) === 'r') &&
                              (item['Detalhe da Atividade'].charAt(2) === 'E' || item['Detalhe da Atividade'].charAt(2) === 'e') &&
                              (item['Detalhe da Atividade'].charAt(3) === 'V' || item['Detalhe da Atividade'].charAt(3) === 'v')
                         )
                              return item
                    })



                    // ANTES DE RETIRAR PREV

                    const data = rowObject

                    // FUNÇAO PARA RETIRAR PREVENTIVA =================

                    function removeItem(arr, prop, value) {
                         prop.toUpperCase()
                         return arr.filter(function (i) { return i[prop] !== value })
                    }


                    rowObject = removeItem(rowObject, 'Detalhe da Atividade', 'PREV/BANDA LARGA')
                    rowObject = removeItem(rowObject, 'Detalhe da Atividade', 'PREV/PÓS CONTATO')
                    rowObject = removeItem(rowObject, 'Detalhe da Atividade', 'Prev/Pós Contato')


                    //===================== DADOS CIDADES                   
                    let dataArc = rowObject.filter(item => item.Cidade === 'ARACRUZ')
                    let dataCim = rowObject.filter(item => item.Cidade === 'CACHOEIRO DE ITAPEMIRIM')
                    let dataCca = rowObject.filter(item => item.Cidade === 'CARIACICA')
                    let dataCna = rowObject.filter(item => item.Cidade === 'COLATINA')
                    let dataGri = rowObject.filter(item => item.Cidade === 'GUARAPARI')
                    let dataLns = rowObject.filter(item => item.Cidade === 'LINHARES')
                    let dataSmj = rowObject.filter(item => item.Cidade === 'SANTA MARIA DE JETIBA')
                    let dataSmt = rowObject.filter(item => item.Cidade === 'SAO MATEUS')
                    let dataSea = rowObject.filter(item => item.Cidade === 'SERRA')
                    let dataVva = rowObject.filter(item => item.Cidade === 'VILA VELHA')
                    let dataVta = rowObject.filter(item => item.Cidade === 'VITORIA')
                    let dataVia = rowObject.filter(item => item.Cidade === 'VIANA')




                    //  DADOS CIDADE METALICO =========================
                    const dataMetalico = rowObject.filter(metalico)
                    const metalicoCca = dataCca.filter(metalico)
                    const metalicoCna = dataCna.filter(metalico)
                    const metalicoLns = dataLns.filter(metalico)
                    const metalicoSea = dataSea.filter(metalico)
                    const metalicoVva = dataVva.filter(metalico)
                    const metalicoVta = dataVta.filter(metalico)


                    // ============= DADOS POR CIDADE GPON 

                    const dataGpon = rowObject.filter(gpon)
                    const gponArc = dataArc.filter(gpon)
                    const gponCim = dataCim.filter(gpon)
                    const gponCca = dataCca.filter(gpon)
                    const gponCna = dataCna.filter(gpon)
                    const gponGri = dataGri.filter(gpon)
                    const gponLns = dataLns.filter(gpon)
                    const gponSmj = dataSmj.filter(gpon)
                    const gponSmt = dataSmt.filter(gpon)
                    const gponSea = dataSea.filter(gpon)
                    const gponVva = dataVva.filter(gpon)
                    const gponVta = dataVta.filter(gpon)
                    const gponVia = dataVia.filter(gpon)


                    // ====CREATE CABEÇALHO TABELA PRODUÇAO GPON ====
                    // const titleProducao = document.createElement('h1')
                    // titleProducao.className = 'title-producao'
                    // titleProducao.innerHTML = 'PRODUÇÃO'
                    // const tProducao = document.getElementById("producao")
                    // tProducao.appendChild(titleProducao)


                    const titleGpon = document.createElement('span')
                    titleGpon.innerHTML = 'GPON'
                    const tGpon = document.getElementById("title-gpon")
                    tGpon.append(titleGpon)

                    const tdCidade = document.createElement('td')
                    tdCidade.className = 'tdCidade'
                    tdCidade.innerHTML = 'CIDADE'

                    const tdConcluida = document.createElement('td')
                    tdConcluida.className = 'tdConcluida'
                    tdConcluida.innerHTML = 'CONCLUIDA'

                    const tdIniciada = document.createElement('td')
                    tdIniciada.className = 'tdIniciada'
                    tdIniciada.innerHTML = 'INICIADA'

                    const tdNaoiniciada = document.createElement('td')
                    tdNaoiniciada.className = 'tdNin'
                    tdNaoiniciada.innerHTML = 'NÃO INICIADA'

                    const total = document.createElement('td')
                    total.className = 'tdTotal'
                    total.innerHTML = 'TOTAL'

                    const tabela = document.getElementById('cabecalho')
                    tabela.append(tdCidade)
                    tabela.append(tdConcluida)
                    tabela.append(tdIniciada)
                    tabela.append(tdNaoiniciada)
                    tabela.append(total)

                    // ======== CREATE CIDADE GPON ==


                    // ARACRUZ
                    const tdArcGpon = document.createElement('td')
                    tdArcGpon.innerHTML = 'ARACRUZ'
                    const colArcGpon = document.getElementById('arc')
                    colArcGpon.append(tdArcGpon)


                    const tdConArcGpon = document.createElement('td')
                    tdConArcGpon.innerHTML = gponArc.filter(statusConcluida).length
                    const colConArcGpon = document.getElementById('arc')
                    colConArcGpon.append(tdConArcGpon)

                    const tdIniArcGpon = document.createElement('td')
                    tdIniArcGpon.innerHTML = gponArc.filter(statusIniciada).length
                    const conIniArcGpon = document.getElementById('arc')
                    conIniArcGpon.append(tdIniArcGpon)

                    const tdNinArcGpon = document.createElement('td')
                    tdNinArcGpon.innerHTML = gponArc.filter(statusNaoIniciada).length
                    const colNinArcGpon = document.getElementById('arc')
                    colNinArcGpon.append(tdNinArcGpon)

                    const tdTotalArcGpon = document.createElement('td')
                    tdTotalArcGpon.innerHTML = gponArc.filter(conIniNin).length
                    const colTotalArcGpon = document.getElementById('arc')
                    colTotalArcGpon.append(tdTotalArcGpon)


                    // CACHOEIRO
                    const tdCimGpon = document.createElement('td')
                    tdCimGpon.innerHTML = 'CACHOEIRO'
                    const colCimGpon = document.getElementById('cim')
                    colCimGpon.append(tdCimGpon)


                    const tdConCimGpon = document.createElement('td')
                    tdConCimGpon.innerHTML = gponCim.filter(statusConcluida).length
                    const colConCimGpon = document.getElementById('cim')
                    colConCimGpon.append(tdConCimGpon)

                    const tdIniCimGpon = document.createElement('td')
                    tdIniCimGpon.innerHTML = gponCim.filter(statusIniciada).length
                    const colIniCimGpon = document.getElementById('cim')
                    colIniCimGpon.append(tdIniCimGpon)

                    const tdNinCimGPon = document.createElement('td')
                    tdNinCimGPon.innerHTML = gponCim.filter(statusNaoIniciada).length
                    const colNinCimGpon = document.getElementById('cim')
                    colNinCimGpon.append(tdNinCimGPon)

                    const tdTotalCimGpon = document.createElement('td')
                    tdTotalCimGpon.innerHTML = gponCim.filter(conIniNin).length
                    const colTotalCimGpon = document.getElementById('cim')
                    colTotalCimGpon.append(tdTotalCimGpon)


                    // CARIACICA
                    const tdCcaGpon = document.createElement('td')
                    tdCcaGpon.innerHTML = 'CARIACICA'
                    const colCcaGpon = document.getElementById('cca')
                    colCcaGpon.append(tdCcaGpon)


                    const tdConCcaGpon = document.createElement('td')
                    tdConCcaGpon.innerHTML = gponCca.filter(statusConcluida).length
                    const colConCcaGpon = document.getElementById('cca')
                    colConCcaGpon.append(tdConCcaGpon)

                    const tdIniCcaGpon = document.createElement('td')
                    tdIniCcaGpon.innerHTML = gponCca.filter(statusIniciada).length
                    const colIniCcaGpon = document.getElementById('cca')
                    colIniCcaGpon.append(tdIniCcaGpon)

                    const tdNinCcaGpon = document.createElement('td')
                    tdNinCcaGpon.innerHTML = gponCca.filter(statusNaoIniciada).length
                    const colNinCcaGpon = document.getElementById('cca')
                    colNinCcaGpon.append(tdNinCcaGpon)

                    const tdTotalCcaGpon = document.createElement('td')
                    tdTotalCcaGpon.innerHTML = gponCca.filter(conIniNin).length
                    const colTotalCcaGpon = document.getElementById('cca')
                    colTotalCcaGpon.append(tdTotalCcaGpon)


                    // COLATINA
                    const tdCnaGpon = document.createElement('td')
                    tdCnaGpon.innerHTML = 'COLATINA'
                    const colCnaGpon = document.getElementById('cna')
                    colCnaGpon.append(tdCnaGpon)

                    const tdConCnaGpon = document.createElement('td')
                    tdConCnaGpon.innerHTML = gponCna.filter(statusConcluida).length
                    const colConCnaGpon = document.getElementById('cna')
                    colConCnaGpon.append(tdConCnaGpon)

                    const tdIniCnaGpon = document.createElement('td')
                    tdIniCnaGpon.innerHTML = gponCna.filter(statusIniciada).length
                    const colIniCnaGpon = document.getElementById('cna')
                    colIniCnaGpon.append(tdIniCnaGpon)

                    const tdNinCnaGpon = document.createElement('td')
                    tdNinCnaGpon.innerHTML = gponCna.filter(statusNaoIniciada).length
                    const colNinCnaGpon = document.getElementById('cna')
                    colNinCnaGpon.append(tdNinCnaGpon)

                    const tdTotalCnaGpon = document.createElement('td')
                    tdTotalCnaGpon.innerHTML = gponCna.filter(conIniNin).length
                    const colTotalCnaGpon = document.getElementById('cna')
                    colTotalCnaGpon.append(tdTotalCnaGpon)


                    // GUARAPARI
                    const tdGriGpon = document.createElement('td')
                    tdGriGpon.innerHTML = 'GUARAPARI'
                    const colGriGpon = document.getElementById('gri')
                    colGriGpon.append(tdGriGpon)

                    const tdConGriGpon = document.createElement('td')
                    tdConGriGpon.innerHTML = gponGri.filter(statusConcluida).length
                    const colConGriGpon = document.getElementById('gri')
                    colConGriGpon.append(tdConGriGpon)

                    const tdIniGriGpon = document.createElement('td')
                    tdIniGriGpon.innerHTML = gponGri.filter(statusIniciada).length
                    const colIniGriGpon = document.getElementById('gri')
                    colIniGriGpon.append(tdIniGriGpon)

                    const tdNinGriGpon = document.createElement('td')
                    tdNinGriGpon.innerHTML = gponGri.filter(statusNaoIniciada).length
                    const colNinGriGpon = document.getElementById('gri')
                    colNinGriGpon.append(tdNinGriGpon)

                    const tdTotalGriGpon = document.createElement('td')
                    tdTotalGriGpon.innerHTML = gponGri.filter(conIniNin).length
                    const colTotalGriGpon = document.getElementById('gri')
                    colTotalGriGpon.append(tdTotalGriGpon)


                    // LINHARES
                    const tdLnsGpon = document.createElement('td')
                    tdLnsGpon.innerHTML = 'LINHARES'
                    const colLnsGpon = document.getElementById('lns')
                    colLnsGpon.append(tdLnsGpon)

                    const tdConLnsGpon = document.createElement('td')
                    tdConLnsGpon.innerHTML = gponLns.filter(statusConcluida).length
                    const colConLnsGpon = document.getElementById('lns')
                    colConLnsGpon.append(tdConLnsGpon)

                    const tdIniLnsGpon = document.createElement('td')
                    tdIniLnsGpon.innerHTML = gponLns.filter(statusIniciada).length
                    const colIniLnsGpon = document.getElementById('lns')
                    colIniLnsGpon.append(tdIniLnsGpon)

                    const tdNinLnsGpon = document.createElement('td')
                    tdNinLnsGpon.innerHTML = gponLns.filter(statusNaoIniciada).length
                    const colNinLnsGpon = document.getElementById('lns')
                    colNinLnsGpon.append(tdNinLnsGpon)

                    const tdTotalLnsGpon = document.createElement('td')
                    tdTotalLnsGpon.innerHTML = gponLns.filter(conIniNin).length
                    const colTotalLnsGpon = document.getElementById('lns')
                    colTotalLnsGpon.append(tdTotalLnsGpon)



                    // SANTA MARIA DE JETIBÁ

                    const tdSmjGpon = document.createElement('td')
                    tdSmjGpon.innerHTML = 'SANTA MARIA'
                    const colSmjGpon = document.getElementById('smj')
                    colSmjGpon.append(tdSmjGpon)

                    const tdConSmjGpon = document.createElement('td')
                    tdConSmjGpon.innerHTML = gponSmj.filter(statusConcluida).length
                    const colConSmjGpon = document.getElementById('smj')
                    colConSmjGpon.append(tdConSmjGpon)

                    const tdIniSmjGpon = document.createElement('td')
                    tdIniSmjGpon.innerHTML = gponSmj.filter(statusIniciada).length
                    const colIniSmjGpon = document.getElementById('smj')
                    colIniSmjGpon.append(tdIniSmjGpon)

                    const tdNinSmjGpon = document.createElement('td')
                    tdNinSmjGpon.innerHTML = gponSmj.filter(statusNaoIniciada).length
                    const colNinSmjGpon = document.getElementById('smj')
                    colNinSmjGpon.append(tdNinSmjGpon)

                    const tdTotalSmjGpon = document.createElement('td')
                    tdTotalSmjGpon.innerHTML = gponSmj.filter(conIniNin).length
                    const colTotalSmjGpon = document.getElementById('smj')
                    colTotalSmjGpon.append(tdTotalSmjGpon)



                    // SÃO MATEUS
                    const tdSmtGpon = document.createElement('td')
                    tdSmtGpon.innerHTML = 'SÃO MATEUS'
                    const colSmtGpon = document.getElementById('smt')
                    colSmtGpon.append(tdSmtGpon)

                    const tdConSmtGpon = document.createElement('td')
                    tdConSmtGpon.innerHTML = gponSmt.filter(statusConcluida).length
                    const colConSmtGpon = document.getElementById('smt')
                    colConSmtGpon.append(tdConSmtGpon)

                    const tdIniSmtGpon = document.createElement('td')
                    tdIniSmtGpon.innerHTML = gponSmt.filter(statusIniciada).length
                    const colIniSmtGpon = document.getElementById('smt')
                    colIniSmtGpon.append(tdIniSmtGpon)

                    const tdNinSmtGpon = document.createElement('td')
                    tdNinSmtGpon.innerHTML = gponSmt.filter(statusNaoIniciada).length
                    const colNinSmtGpon = document.getElementById('smt')
                    colNinSmtGpon.append(tdNinSmtGpon)

                    const tdTotalSmtGpon = document.createElement('td')
                    tdTotalSmtGpon.innerHTML = gponSmt.filter(conIniNin).length
                    const colTotalSmtGpon = document.getElementById('smt')
                    colTotalSmtGpon.append(tdTotalSmtGpon)

                    // SERRA
                    const tdSeaGpon = document.createElement('td')
                    tdSeaGpon.innerHTML = 'SERRA'
                    const colSeaGpon = document.getElementById('sea')
                    colSeaGpon.append(tdSeaGpon)


                    const tdConSeaGpon = document.createElement('td')
                    tdConSeaGpon.innerHTML = gponSea.filter(statusConcluida).length
                    const colConSeaGpon = document.getElementById('sea')
                    colConSeaGpon.append(tdConSeaGpon)

                    const tdIniSeaGpon = document.createElement('td')
                    tdIniSeaGpon.innerHTML = gponSea.filter(statusIniciada).length
                    const colIniSeaGpon = document.getElementById('sea')
                    colIniSeaGpon.append(tdIniSeaGpon)

                    const tdNinSeaGpon = document.createElement('td')
                    tdNinSeaGpon.innerHTML = gponSea.filter(statusNaoIniciada).length
                    const colNinSeaGpon = document.getElementById('sea')
                    colNinSeaGpon.append(tdNinSeaGpon)

                    const tdTotalSea = document.createElement('td')
                    tdTotalSea.innerHTML = gponSea.filter(conIniNin).length
                    const colTotalSea = document.getElementById('sea')
                    colTotalSea.append(tdTotalSea)


                    // VILA VELHA
                    const tdVvaGpon = document.createElement('td')
                    tdVvaGpon.innerHTML = 'VILA VELHA'
                    const colVvaGpon = document.getElementById('vva')
                    colVvaGpon.append(tdVvaGpon)

                    const tdConVvaGpon = document.createElement('td')
                    tdConVvaGpon.innerHTML = gponVva.filter(statusConcluida).length
                    const colConVva = document.getElementById('vva')
                    colConVva.append(tdConVvaGpon)

                    const tdIniVvaGpon = document.createElement('td')
                    tdIniVvaGpon.innerHTML = gponVva.filter(statusIniciada).length
                    const colIniVva = document.getElementById('vva')
                    colIniVva.append(tdIniVvaGpon)

                    const tdNinVvaGpon = document.createElement('td')
                    tdNinVvaGpon.innerHTML = gponVva.filter(statusNaoIniciada).length
                    const colNinVvaGpon = document.getElementById('vva')
                    colNinVvaGpon.append(tdNinVvaGpon)

                    const tdTotalVvaGpon = document.createElement('td')
                    tdTotalVvaGpon.innerHTML = gponVva.filter(conIniNin).length
                    const colTotalVvaGpon = document.getElementById('vva')
                    colTotalVvaGpon.append(tdTotalVvaGpon)


                    // VITORIA
                    const tdVtaGpon = document.createElement('td')
                    tdVtaGpon.innerHTML = 'VITÓRIA'
                    const colVtaGpon = document.getElementById('vta')
                    colVtaGpon.append(tdVtaGpon)

                    const tdConVtaGpon = document.createElement('td')
                    tdConVtaGpon.innerHTML = gponVta.filter(statusConcluida).length
                    const colConVtaGpon = document.getElementById('vta')
                    colConVtaGpon.append(tdConVtaGpon)

                    const tdIniVtaGpon = document.createElement('td')
                    tdIniVtaGpon.innerHTML = gponVta.filter(statusIniciada).length
                    const colIniVtaGpon = document.getElementById('vta')
                    colIniVtaGpon.append(tdIniVtaGpon)

                    const tdNinVtaGpon = document.createElement('td')
                    tdNinVtaGpon.innerHTML = gponVta.filter(statusNaoIniciada).length
                    const colNinVtaGpon = document.getElementById('vta')
                    colNinVtaGpon.append(tdNinVtaGpon)

                    const tdTotalVta = document.createElement('td')
                    tdTotalVta.innerHTML = gponVta.filter(conIniNin).length
                    const colTotalVta = document.getElementById('vta')
                    colTotalVta.append(tdTotalVta)


                    // VIANA
                    const tdViaGpon = document.createElement('td')
                    tdViaGpon.innerHTML = 'VIANA'
                    const colViaGpon = document.getElementById('via')
                    colViaGpon.append(tdViaGpon)

                    const tdConViaGpon = document.createElement('td')
                    tdConViaGpon.innerHTML = gponVia.filter(statusConcluida).length
                    const colConViaGpon = document.getElementById('via')
                    colConViaGpon.append(tdConViaGpon)

                    const tdIniViaGpon = document.createElement('td')
                    tdIniViaGpon.innerHTML = gponVia.filter(statusIniciada).length
                    const colIniViaGpon = document.getElementById('via')
                    colIniViaGpon.append(tdIniViaGpon)

                    const tdNinViaGpon = document.createElement('td')
                    tdNinViaGpon.innerHTML = gponVia.filter(statusNaoIniciada).length
                    const colNinViaGpon = document.getElementById('via')
                    colNinViaGpon.append(tdNinViaGpon)

                    const tdTotalVia = document.createElement('td')
                    tdTotalVia.innerHTML = gponVia.filter(conIniNin).length
                    const colTotalVia = document.getElementById('via')
                    colTotalVia.append(tdTotalVia)






                    // TOTAL
                    const tdGpon = document.createElement('td')
                    tdGpon.className = 'tdGpon'
                    tdGpon.innerHTML = 'TOTAL'
                    const colGpon = document.getElementById('total')
                    colGpon.append(tdGpon)

                    const tdConGpon = document.createElement('td')
                    tdConGpon.innerHTML = dataGpon.filter(statusConcluida).length
                    const colConGpon = document.getElementById('total')
                    colConGpon.append(tdConGpon)

                    const tdIniGpon = document.createElement('td')
                    tdIniGpon.innerHTML = dataGpon.filter(statusIniciada).length
                    const colIniGpon = document.getElementById('total')
                    colIniGpon.append(tdIniGpon)

                    const tdNinGpon = document.createElement('td')
                    tdNinGpon.innerHTML = dataGpon.filter(statusNaoIniciada).length
                    const colNinGpon = document.getElementById('total')
                    colNinGpon.append(tdNinGpon)

                    const tdSumGpon = document.createElement('td')
                    tdSumGpon.innerHTML = dataGpon.filter(conIniNin).length
                    const colSumGpon = document.getElementById('total')
                    colSumGpon.append(tdSumGpon)


                    //============ CREATE CIDADADE METALICO                 


                    const titleMetalico = document.createElement('span')
                    titleMetalico.innerHTML = 'METALICO'
                    const tMetalico = document.getElementById("title-metalico")
                    tMetalico.append(titleMetalico)

                    const tdCidadeMetalico = document.createElement('td')
                    tdCidadeMetalico.className = 'tdCidade'
                    tdCidadeMetalico.innerHTML = 'CIDADE'

                    const tdConcluidaMetalico = document.createElement('td')
                    tdConcluidaMetalico.className = 'tdConcluida'
                    tdConcluidaMetalico.innerHTML = 'CONCLUIDA'

                    const tdIniciadaMetalico = document.createElement('td')
                    tdIniciadaMetalico.className = 'tdIniciada'
                    tdIniciadaMetalico.innerHTML = 'INICIADA'
                    const tdNaoiniciadaMetalico = document.createElement('td')

                    tdNaoiniciadaMetalico.innerHTML = 'NÃO INICIADA'
                    tdNaoiniciadaMetalico.className = 'tdNin'

                    const totalMetalico = document.createElement('td')
                    totalMetalico.className = 'tdTotal'
                    totalMetalico.innerHTML = 'TOTAL'

                    const tabelaMetalico = document.getElementById('cabecalho-metalico')
                    tabelaMetalico.append(tdCidadeMetalico)
                    tabelaMetalico.append(tdConcluidaMetalico)
                    tabelaMetalico.append(tdIniciadaMetalico)
                    tabelaMetalico.append(tdNaoiniciadaMetalico)
                    tabelaMetalico.append(totalMetalico)

                    // DADOS CIDADE METALICO 

                    // CARIACICA
                    const tdCcaMetalico = document.createElement('td')
                    tdCcaMetalico.innerHTML = 'CARIACICA'
                    const colCcaMetalico = document.getElementById('cca-metalico')
                    colCcaMetalico.append(tdCcaMetalico)

                    const tdConCcaMetalico = document.createElement('td')
                    tdConCcaMetalico.innerHTML = metalicoCca.filter(statusConcluida).length
                    const colConCcaMetalico = document.getElementById('cca-metalico')
                    colConCcaMetalico.append(tdConCcaMetalico)

                    const tdIniCcaMetalico = document.createElement('td')
                    tdIniCcaMetalico.innerHTML = metalicoCca.filter(statusIniciada).length
                    const colIniCcaMetalico = document.getElementById('cca-metalico')
                    colIniCcaMetalico.append(tdIniCcaMetalico)

                    const tdNinCcaMetalico = document.createElement('td')
                    tdNinCcaMetalico.innerHTML = metalicoCca.filter(statusNaoIniciada).length
                    const colNinCcaMetalico = document.getElementById('cca-metalico')
                    colNinCcaMetalico.append(tdNinCcaMetalico)

                    const tdTotalCcaMetalico = document.createElement('td')
                    tdTotalCcaMetalico.innerHTML = metalicoCca.filter(conIniNin).length
                    const colTotalCcaMetalico = document.getElementById('cca-metalico')
                    colTotalCcaMetalico.append(tdTotalCcaMetalico)


                    // COLATINA

                    const tdCnaMetalico = document.createElement('td')
                    tdCnaMetalico.innerHTML = 'COLATINA'
                    const colCnaMetalico = document.getElementById('cna-metalico')
                    colCnaMetalico.append(tdCnaMetalico)

                    const tdConCnaMetalico = document.createElement('td')
                    tdConCnaMetalico.innerHTML = metalicoCna.filter(statusConcluida).length
                    const colConCnaMetalico = document.getElementById('cna-metalico')
                    colConCnaMetalico.append(tdConCnaMetalico)

                    const tdIniCnaMetalico = document.createElement('td')
                    tdIniCnaMetalico.innerHTML = metalicoCna.filter(statusIniciada).length
                    const colIniCnaMetalico = document.getElementById('cna-metalico')
                    colIniCnaMetalico.append(tdIniCnaMetalico)

                    const tdNinCnaMetalico = document.createElement('td')
                    tdNinCnaMetalico.innerHTML = metalicoCna.filter(statusNaoIniciada).length
                    const colNinCnaMetalico = document.getElementById('cna-metalico')
                    colNinCnaMetalico.append(tdNinCnaMetalico)

                    const tdTotalCnaMetalico = document.createElement('td')
                    tdTotalCnaMetalico.innerHTML = metalicoCna.filter(conIniNin).length
                    const colTotalCnaMetalico = document.getElementById('cna-metalico')
                    colTotalCnaMetalico.append(tdTotalCnaMetalico)


                    // LINHARES

                    const tdLnsMetalico = document.createElement('td')
                    tdLnsMetalico.innerHTML = 'LINHARES'
                    const colLnsMetalico = document.getElementById('lns-metalico')
                    colLnsMetalico.append(tdLnsMetalico)

                    const tdConLnsMetalico = document.createElement('td')
                    tdConLnsMetalico.innerHTML = metalicoLns.filter(statusConcluida).length
                    const colConLnsMetalico = document.getElementById('lns-metalico')
                    colConLnsMetalico.append(tdConLnsMetalico)

                    const tdIniLnsMetalico = document.createElement('td')
                    tdIniLnsMetalico.innerHTML = metalicoLns.filter(statusIniciada).length
                    const colIniLnsMetalico = document.getElementById('lns-metalico')
                    colIniLnsMetalico.append(tdIniLnsMetalico)

                    const tdNinLnsMetalico = document.createElement('td')
                    tdNinLnsMetalico.innerHTML = metalicoLns.filter(statusNaoIniciada).length
                    const colNinLnsMetalico = document.getElementById('lns-metalico')
                    colNinLnsMetalico.append(tdNinLnsMetalico)

                    const tdTotalLnsMetalico = document.createElement('td')
                    tdTotalLnsMetalico.innerHTML = metalicoLns.filter(conIniNin).length
                    const colTotalLnsMetalico = document.getElementById('lns-metalico')
                    colTotalLnsMetalico.append(tdTotalLnsMetalico)


                    // SERRA

                    const tdSeaMetalico = document.createElement('td')
                    tdSeaMetalico.innerHTML = 'SERRA'
                    const colSeaMetalico = document.getElementById('sea-metalico')
                    colSeaMetalico.append(tdSeaMetalico)

                    const tdConSeaMetalico = document.createElement('td')
                    tdConSeaMetalico.innerHTML = metalicoSea.filter(statusConcluida).length
                    const colConSeaMetalico = document.getElementById('sea-metalico')
                    colConSeaMetalico.append(tdConSeaMetalico)

                    const tdIniSeaMetalico = document.createElement('td')
                    tdIniSeaMetalico.innerHTML = metalicoSea.filter(statusIniciada).length
                    const colIniSeaMetalico = document.getElementById('sea-metalico')
                    colIniSeaMetalico.append(tdIniSeaMetalico)

                    const tdNinSeaMetalico = document.createElement('td')
                    tdNinSeaMetalico.innerHTML = metalicoSea.filter(statusNaoIniciada).length
                    const colNinSeaMetalico = document.getElementById('sea-metalico')
                    colNinSeaMetalico.append(tdNinSeaMetalico)

                    const tdTotalSeaMetalico = document.createElement('td')
                    tdTotalSeaMetalico.innerHTML = metalicoSea.filter(conIniNin).length
                    const colTotalSeaMetalico = document.getElementById('sea-metalico')
                    colTotalSeaMetalico.append(tdTotalSeaMetalico)


                    // VILA VELHA

                    const tdVvaMetalico = document.createElement('td')
                    tdVvaMetalico.innerHTML = 'VILA VELHA'
                    const colVvaMetalico = document.getElementById('vva-metalico')
                    colVvaMetalico.append(tdVvaMetalico)

                    const tdConVvaMetalico = document.createElement('td')
                    tdConVvaMetalico.innerHTML = metalicoVva.filter(statusConcluida).length
                    const colConVvaMetalico = document.getElementById('vva-metalico')
                    colConVvaMetalico.append(tdConVvaMetalico)

                    const tdIniVvaMetalico = document.createElement('td')
                    tdIniVvaMetalico.innerHTML = metalicoVva.filter(statusIniciada).length
                    const colIniVvaMetalico = document.getElementById('vva-metalico')
                    colIniVvaMetalico.append(tdIniVvaMetalico)

                    const tdNinVvaMetalico = document.createElement('td')
                    tdNinVvaMetalico.innerHTML = metalicoVva.filter(statusNaoIniciada).length
                    const colNinVvaMetalico = document.getElementById('vva-metalico')
                    colNinVvaMetalico.append(tdNinVvaMetalico)

                    const tdTotalVvaMetalico = document.createElement('td')
                    tdTotalVvaMetalico.innerHTML = metalicoVva.filter(conIniNin).length
                    const colTotalVvaMetalico = document.getElementById('vva-metalico')
                    colTotalVvaMetalico.append(tdTotalVvaMetalico)

                    // VITÓRIA

                    const tdVtaMetalico = document.createElement('td')
                    tdVtaMetalico.innerHTML = 'VITÓRIA'
                    const colVtaMetalico = document.getElementById('vta-metalico')
                    colVtaMetalico.append(tdVtaMetalico)

                    const tdConVtaMetalico = document.createElement('td')
                    tdConVtaMetalico.innerHTML = metalicoVta.filter(statusConcluida).length
                    const colConVtaMetalico = document.getElementById('vta-metalico')
                    colConVtaMetalico.append(tdConVtaMetalico)

                    const tdIniVtaMetalico = document.createElement('td')
                    tdIniVtaMetalico.innerHTML = metalicoVta.filter(statusIniciada).length
                    const colIniVtaMetalico = document.getElementById('vta-metalico')
                    colIniVtaMetalico.append(tdIniVtaMetalico)

                    const tdNinVtaMetalico = document.createElement('td')
                    tdNinVtaMetalico.innerHTML = metalicoVta.filter(statusNaoIniciada).length
                    const colNinVtaMetalico = document.getElementById('vta-metalico')
                    colNinVtaMetalico.append(tdNinVtaMetalico)

                    const tdTotalVtaMetalico = document.createElement('td')
                    tdTotalVtaMetalico.innerHTML = metalicoVta.filter(conIniNin).length
                    const colTotalVtaMetalico = document.getElementById('vta-metalico')
                    colTotalVtaMetalico.append(tdTotalVtaMetalico)


                    // ========== TOTAL METALICO 

                    const tdMetalico = document.createElement('td')
                    tdMetalico.innerHTML = 'TOTAL'
                    const colMetalico = document.getElementById('total-metalico')
                    colMetalico.append(tdMetalico)

                    const tdConMetalico = document.createElement('td')
                    tdConMetalico.innerHTML = dataMetalico.filter(statusConcluida).length
                    const colConMetalico = document.getElementById('total-metalico')
                    colConMetalico.append(tdConMetalico)

                    const tdIniMetalico = document.createElement('td')
                    tdIniMetalico.innerHTML = dataMetalico.filter(statusIniciada).length
                    const colIniMetalico = document.getElementById('total-metalico')
                    colIniMetalico.append(tdIniMetalico)

                    const tdNinMetalico = document.createElement('td')
                    tdNinMetalico.innerHTML = dataMetalico.filter(statusNaoIniciada).length
                    const colNinMetalico = document.getElementById('total-metalico')
                    colNinMetalico.append(tdNinMetalico)

                    const tdSumMetalico = document.createElement('td')
                    tdSumMetalico.innerHTML = dataMetalico.filter(conIniNin).length
                    const colSumMetalico = document.getElementById('total-metalico')
                    colSumMetalico.append(tdSumMetalico)






                    // ====CREATE CABEÇALHO TABELA PRODUÇAO PREVENTIVA ====


                    const titlePrev = document.createElement('span')
                    titlePrev.innerHTML = 'PREVENTIVA'
                    const tPrev = document.getElementById("title-preventiva")
                    tPrev.append(titlePrev)

                    const tdCidadePrev = document.createElement('td')
                    tdCidadePrev.className = 'tdCidade'
                    tdCidadePrev.innerHTML = 'CIDADE'

                    const tdConcluidaPrev = document.createElement('td')
                    tdConcluidaPrev.className = 'tdConcluida'
                    tdConcluidaPrev.innerHTML = 'CONCLUIDA'

                    const tdIniciadaPrev = document.createElement('td')
                    tdIniciadaPrev.className = 'tdIniciada'
                    tdIniciadaPrev.innerHTML = 'INICIADA'

                    const tdNaoiniciadaPrev = document.createElement('td')
                    tdNaoiniciadaPrev.className = 'tdNin'
                    tdNaoiniciadaPrev.innerHTML = 'NÃO INICIADA'

                    const totalPrev = document.createElement('td')
                    totalPrev.className = 'tdTotal'
                    totalPrev.innerHTML = 'TOTAL'

                    const tabelaPrev = document.getElementById('cabecalho-preventiva')
                    tabelaPrev.append(tdCidadePrev)
                    tabelaPrev.append(tdConcluidaPrev)
                    tabelaPrev.append(tdIniciadaPrev)
                    tabelaPrev.append(tdNaoiniciadaPrev)
                    tabelaPrev.append(totalPrev)


                    // ============= DADOS PREVENTIVA

                    const prevArc = data.filter(item => item.Cidade === 'ARACRUZ').filter(prev)
                    const prevCim = data.filter(item => item.Cidade === 'CACHOEIRO DE ITAPEMIRIM').filter(prev)
                    const prevCca = data.filter(item => item.Cidade === 'CARIACICA').filter(prev)
                    const prevCna = data.filter(item => item.Cidade === 'COLATINA').filter(prev)
                    const prevGri = data.filter(item => item.Cidade === 'GUARAPARI').filter(prev)
                    const prevLns = data.filter(item => item.Cidade === 'LINHARES').filter(prev)
                    const prevSmj = data.filter(item => item.Cidade === 'SANTA MARIA DE JETIBA').filter(prev)
                    const prevSmt = data.filter(item => item.Cidade === 'SAO MATEUS').filter(prev)
                    const prevSea = data.filter(item => item.Cidade === 'SERRA').filter(prev)
                    const prevVia = data.filter(item => item.Cidade === 'VIANA').filter(prev)
                    const prevVva = data.filter(item => item.Cidade === 'VILA VELHA').filter(prev)
                    const prevVta = data.filter(item => item.Cidade === 'VITORIA').filter(prev)

                    // ======== PRODUÇAO PREV ====


                    // ARACRUZ
                    const tdArcPrev = document.createElement('td')
                    tdArcPrev.innerHTML = 'ARACRUZ'
                    const colArcGponPrev = document.getElementById('arc-preventiva')
                    colArcGponPrev.append(tdArcPrev)


                    const tdConArcPrev = document.createElement('td')
                    tdConArcPrev.innerHTML = prevArc.filter(statusConcluida).length
                    const colConArcPrev = document.getElementById('arc-preventiva')
                    colConArcPrev.append(tdConArcPrev)

                    const tdIniArcPrev = document.createElement('td')
                    tdIniArcPrev.innerHTML = prevArc.filter(statusIniciada).length
                    const conIniArcPrev = document.getElementById('arc-preventiva')
                    conIniArcPrev.append(tdIniArcPrev)

                    const tdNinArcPrev = document.createElement('td')
                    tdNinArcPrev.innerHTML = prevArc.filter(statusNaoIniciada).length
                    const colNinArcPrev = document.getElementById('arc-preventiva')
                    colNinArcPrev.append(tdNinArcPrev)

                    const tdTotalArcPrev = document.createElement('td')
                    tdTotalArcPrev.innerHTML = prevArc.filter(conIniNin).length
                    const colTotalArcPrev = document.getElementById('arc-preventiva')
                    colTotalArcPrev.append(tdTotalArcPrev)



                    // CACHOEIRO

                    const tdCimPrev = document.createElement('td')
                    tdCimPrev.innerHTML = 'CACHOEIRO'
                    const colCimPrev = document.getElementById('cim-preventiva')
                    colCimPrev.append(tdCimPrev)


                    const tdConCimPrev = document.createElement('td')
                    tdConCimPrev.innerHTML = prevCim.filter(statusConcluida).length
                    const colConCimPrev = document.getElementById('cim-preventiva')
                    colConCimPrev.append(tdConCimPrev)

                    const tdIniCimPrev = document.createElement('td')
                    tdIniCimPrev.innerHTML = prevCim.filter(statusIniciada).length
                    const colIniCimPrev = document.getElementById('cim-preventiva')
                    colIniCimPrev.append(tdIniCimPrev)

                    const tdNinCimPrev = document.createElement('td')
                    tdNinCimPrev.innerHTML = prevCim.filter(statusNaoIniciada).length
                    const colNinCimPrev = document.getElementById('cim-preventiva')
                    colNinCimPrev.append(tdNinCimPrev)

                    const tdTotalCimPrev = document.createElement('td')
                    tdTotalCimPrev.innerHTML = prevCim.filter(conIniNin).length
                    const colTotalCimPrev = document.getElementById('cim-preventiva')
                    colTotalCimPrev.append(tdTotalCimPrev)


                    // CARIACICA
                    const tdCcaPrev = document.createElement('td')
                    tdCcaPrev.innerHTML = 'CARIACICA'
                    const colCcaPrev = document.getElementById('cca-preventiva')
                    colCcaPrev.append(tdCcaPrev)


                    const tdConCcaPrev = document.createElement('td')
                    tdConCcaPrev.innerHTML = prevCca.filter(statusConcluida).length
                    const colConCcaPrev = document.getElementById('cca-preventiva')
                    colConCcaPrev.append(tdConCcaPrev)

                    const tdIniCcaPrev = document.createElement('td')
                    tdIniCcaPrev.innerHTML = prevCca.filter(statusIniciada).length
                    const colIniCcaPrev = document.getElementById('cca-preventiva')
                    colIniCcaPrev.append(tdIniCcaPrev)

                    const tdNinCcaPrev = document.createElement('td')
                    tdNinCcaPrev.innerHTML = prevCca.filter(statusNaoIniciada).length
                    const colNinCcaPrev = document.getElementById('cca-preventiva')
                    colNinCcaPrev.append(tdNinCcaPrev)

                    const tdTotalCcaPrev = document.createElement('td')
                    tdTotalCcaPrev.innerHTML = prevCca.filter(conIniNin).length
                    const colTotalCcaPrev = document.getElementById('cca-preventiva')
                    colTotalCcaPrev.append(tdTotalCcaPrev)


                    // COLATINA
                    const tdCnaPrev = document.createElement('td')
                    tdCnaPrev.innerHTML = 'COLATINA'
                    const colCnaPrev = document.getElementById('cna-preventiva')
                    colCnaPrev.append(tdCnaPrev)

                    const tdConCnaPrev = document.createElement('td')
                    tdConCnaPrev.innerHTML = prevCna.filter(statusConcluida).length
                    const colConCnaPrev = document.getElementById('cna-preventiva')
                    colConCnaPrev.append(tdConCnaPrev)

                    const tdIniCnaPrev = document.createElement('td')
                    tdIniCnaPrev.innerHTML = prevCna.filter(statusIniciada).length
                    const colIniCnaPrev = document.getElementById('cna-preventiva')
                    colIniCnaPrev.append(tdIniCnaPrev)

                    const tdNinCnaPrev = document.createElement('td')
                    tdNinCnaPrev.innerHTML = prevCna.filter(statusNaoIniciada).length
                    const colNinCnaPrev = document.getElementById('cna-preventiva')
                    colNinCnaPrev.append(tdNinCnaPrev)

                    const tdTotalCnaPrev = document.createElement('td')
                    tdTotalCnaPrev.innerHTML = prevCna.filter(conIniNin).length
                    const colTotalCnaPrev = document.getElementById('cna-preventiva')
                    colTotalCnaPrev.append(tdTotalCnaPrev)


                    // GUARAPARI
                    const tdGriPrev = document.createElement('td')
                    tdGriPrev.innerHTML = 'GUARAPARI'
                    const colGriPrev = document.getElementById('gri-preventiva')
                    colGriPrev.append(tdGriPrev)

                    const tdConGriPrev = document.createElement('td')
                    tdConGriPrev.innerHTML = prevGri.filter(statusConcluida).length
                    const colConGriPrev = document.getElementById('gri-preventiva')
                    colConGriPrev.append(tdConGriPrev)

                    const tdIniGriPrev = document.createElement('td')
                    tdIniGriPrev.innerHTML = prevGri.filter(statusIniciada).length
                    const colIniGriPrev = document.getElementById('gri-preventiva')
                    colIniGriPrev.append(tdIniGriPrev)

                    const tdNinGriPrev = document.createElement('td')
                    tdNinGriPrev.innerHTML = prevGri.filter(statusNaoIniciada).length
                    const colNinGriPrev = document.getElementById('gri-preventiva')
                    colNinGriPrev.append(tdNinGriPrev)

                    const tdTotalGriPrev = document.createElement('td')
                    tdTotalGriPrev.innerHTML = prevGri.filter(conIniNin).length
                    const colTotalGriPrev = document.getElementById('gri-preventiva')
                    colTotalGriPrev.append(tdTotalGriPrev)



                    // LINHARES
                    const tdLnsPrev = document.createElement('td')
                    tdLnsPrev.innerHTML = 'LINHARES'
                    const colLnsPrev = document.getElementById('lns-preventiva')
                    colLnsPrev.append(tdLnsPrev)

                    const tdConLnsPrev = document.createElement('td')
                    tdConLnsPrev.innerHTML = prevLns.filter(statusConcluida).length
                    const colConLnsPrev = document.getElementById('lns-preventiva')
                    colConLnsPrev.append(tdConLnsPrev)

                    const tdIniLnsPrev = document.createElement('td')
                    tdIniLnsPrev.innerHTML = prevLns.filter(statusIniciada).length
                    const colIniLnsPrev = document.getElementById('lns-preventiva')
                    colIniLnsPrev.append(tdIniLnsPrev)

                    const tdNinLnsPrev = document.createElement('td')
                    tdNinLnsPrev.innerHTML = prevLns.filter(statusNaoIniciada).length
                    const colNinLnsPrev = document.getElementById('lns-preventiva')
                    colNinLnsPrev.append(tdNinLnsPrev)

                    const tdTotalLnsPrev = document.createElement('td')
                    tdTotalLnsPrev.innerHTML = prevLns.filter(conIniNin).length
                    const colTotalLnsPrev = document.getElementById('lns-preventiva')
                    colTotalLnsPrev.append(tdTotalLnsPrev)




                    // SANTA MARIA DE JETIBA
                    const tdSmjPrev = document.createElement('td')
                    tdSmjPrev.innerHTML = 'SANTA MARIA'
                    const colSmjPrev = document.getElementById('smj-preventiva')
                    colSmjPrev.append(tdSmjPrev)

                    const tdConSmjPrev = document.createElement('td')
                    tdConSmjPrev.innerHTML = prevSmj.filter(statusConcluida).length
                    const colConSmjPrev = document.getElementById('smj-preventiva')
                    colConSmjPrev.append(tdConSmjPrev)

                    const tdIniSmjPrev = document.createElement('td')
                    tdIniSmjPrev.innerHTML = prevSmj.filter(statusIniciada).length
                    const colIniSmjPrev = document.getElementById('smj-preventiva')
                    colIniSmjPrev.append(tdIniSmjPrev)

                    const tdNinSmjPrev = document.createElement('td')
                    tdNinSmjPrev.innerHTML = prevSmj.filter(statusNaoIniciada).length
                    const colNinSmjPrev = document.getElementById('smj-preventiva')
                    colNinSmjPrev.append(tdNinSmjPrev)

                    const tdTotalSmjPrev = document.createElement('td')
                    tdTotalSmjPrev.innerHTML = prevSmj.filter(conIniNin).length
                    const colTotalSmjPrev = document.getElementById('smj-preventiva')
                    colTotalSmjPrev.append(tdTotalSmjPrev)



                    // SÃO MATEUS
                    const tdSmtPrev = document.createElement('td')
                    tdSmtPrev.innerHTML = 'SÃO MATEUS'
                    const colSmtPrev = document.getElementById('smt-preventiva')
                    colSmtPrev.append(tdSmtPrev)

                    const tdConSmtPrev = document.createElement('td')
                    tdConSmtPrev.innerHTML = prevSmt.filter(statusConcluida).length
                    const colConSmtPrev = document.getElementById('smt-preventiva')
                    colConSmtPrev.append(tdConSmtPrev)

                    const tdIniSmtPrev = document.createElement('td')
                    tdIniSmtPrev.innerHTML = prevSmt.filter(statusIniciada).length
                    const colIniSmtPrev = document.getElementById('smt-preventiva')
                    colIniSmtPrev.append(tdIniSmtPrev)

                    const tdNinSmtPrev = document.createElement('td')
                    tdNinSmtPrev.innerHTML = prevSmt.filter(statusNaoIniciada).length
                    const colNinSmtPrev = document.getElementById('smt-preventiva')
                    colNinSmtPrev.append(tdNinSmtPrev)

                    const tdTotalSmtPrev = document.createElement('td')
                    tdTotalSmtPrev.innerHTML = prevSmt.filter(conIniNin).length
                    const colTotalSmtPrev = document.getElementById('smt-preventiva')
                    colTotalSmtPrev.append(tdTotalSmtPrev)

                    // SERRA
                    const tdSeaPrev = document.createElement('td')
                    tdSeaPrev.innerHTML = 'SERRA'
                    const colSeaPrev = document.getElementById('sea-preventiva')
                    colSeaPrev.append(tdSeaPrev)


                    const tdConSeaPrev = document.createElement('td')
                    tdConSeaPrev.innerHTML = prevSea.filter(statusConcluida).length
                    const colConSeaPrev = document.getElementById('sea-preventiva')
                    colConSeaPrev.append(tdConSeaPrev)

                    const tdIniSeaPrev = document.createElement('td')
                    tdIniSeaPrev.innerHTML = prevSea.filter(statusIniciada).length
                    const colIniSeaPrev = document.getElementById('sea-preventiva')
                    colIniSeaPrev.append(tdIniSeaPrev)

                    const tdNinSeaPrev = document.createElement('td')
                    tdNinSeaPrev.innerHTML = prevSea.filter(statusNaoIniciada).length
                    const colNinSeaPrev = document.getElementById('sea-preventiva')
                    colNinSeaPrev.append(tdNinSeaPrev)

                    const tdTotalSeaPrev = document.createElement('td')
                    tdTotalSeaPrev.innerHTML = prevSea.filter(conIniNin).length
                    const colTotalSeaPrev = document.getElementById('sea-preventiva')
                    colTotalSeaPrev.append(tdTotalSeaPrev)



                    // VIANA
                    const tdViaPrev = document.createElement('td')
                    tdViaPrev.innerHTML = 'VIANA'
                    const colViaPrev = document.getElementById('via-preventiva')
                    colViaPrev.append(tdViaPrev)


                    const tdConViaPrev = document.createElement('td')
                    tdConViaPrev.innerHTML = prevVia.filter(statusConcluida).length
                    const colConViaPrev = document.getElementById('via-preventiva')
                    colConViaPrev.append(tdConViaPrev)

                    const tdIniViaPrev = document.createElement('td')
                    tdIniViaPrev.innerHTML = prevVia.filter(statusIniciada).length
                    const colIniViaPrev = document.getElementById('via-preventiva')
                    colIniViaPrev.append(tdIniViaPrev)

                    const tdNinViaPrev = document.createElement('td')
                    tdNinViaPrev.innerHTML = prevVia.filter(statusNaoIniciada).length
                    const colNinViaPrev = document.getElementById('via-preventiva')
                    colNinViaPrev.append(tdNinViaPrev)

                    const tdTotalViaPrev = document.createElement('td')
                    tdTotalViaPrev.innerHTML = prevVia.filter(conIniNin).length
                    const colTotalViaPrev = document.getElementById('via-preventiva')
                    colTotalViaPrev.append(tdTotalViaPrev)


                    // VILA VELHA
                    const tdVvaPrev = document.createElement('td')
                    tdVvaPrev.innerHTML = 'VILA VELHA'
                    const colVvaPrev = document.getElementById('vva-preventiva')
                    colVvaPrev.append(tdVvaPrev)

                    const tdConVvaPrev = document.createElement('td')
                    tdConVvaPrev.innerHTML = prevVva.filter(statusConcluida).length
                    const colConVvaPrev = document.getElementById('vva-preventiva')
                    colConVvaPrev.append(tdConVvaPrev)

                    const tdIniVvaPrev = document.createElement('td')
                    tdIniVvaPrev.innerHTML = prevVva.filter(statusIniciada).length
                    const colIniVvaPrev = document.getElementById('vva-preventiva')
                    colIniVvaPrev.append(tdIniVvaPrev)

                    const tdNinVvaPrev = document.createElement('td')
                    tdNinVvaPrev.innerHTML = prevVva.filter(statusNaoIniciada).length
                    const colNinVvaPrev = document.getElementById('vva-preventiva')
                    colNinVvaPrev.append(tdNinVvaPrev)

                    const tdTotalVvaPrev = document.createElement('td')
                    tdTotalVvaPrev.innerHTML = prevVva.filter(conIniNin).length
                    const colTotalVvaPrev = document.getElementById('vva-preventiva')
                    colTotalVvaPrev.append(tdTotalVvaPrev)


                    // VITORIA
                    const tdVtaPrev = document.createElement('td')
                    tdVtaPrev.innerHTML = 'VITÓRIA'
                    const colVtaPrev = document.getElementById('vta-preventiva')
                    colVtaPrev.append(tdVtaPrev)

                    const tdConVtaPrev = document.createElement('td')
                    tdConVtaPrev.innerHTML = prevVta.filter(statusConcluida).length
                    const colConVtaPrev = document.getElementById('vta-preventiva')
                    colConVtaPrev.append(tdConVtaPrev)

                    const tdIniVtaPrev = document.createElement('td')
                    tdIniVtaPrev.innerHTML = prevVta.filter(statusIniciada).length
                    const colIniVtaPrev = document.getElementById('vta-preventiva')
                    colIniVtaPrev.append(tdIniVtaPrev)

                    const tdNinVtaPrev = document.createElement('td')
                    tdNinVtaPrev.innerHTML = prevVta.filter(statusNaoIniciada).length
                    const colNinVtaPrev = document.getElementById('vta-preventiva')
                    colNinVtaPrev.append(tdNinVtaPrev)

                    const tdTotalVtaPrev = document.createElement('td')
                    tdTotalVtaPrev.innerHTML = prevVta.filter(conIniNin).length
                    const colTotalVtaPrev = document.getElementById('vta-preventiva')
                    colTotalVtaPrev.append(tdTotalVtaPrev)


                    // TOTAL
                    const dataPrev = data.filter(prev)


                    const tdPrev = document.createElement('td')
                    tdPrev.className = 'tdGpon'
                    tdPrev.innerHTML = 'TOTAL'
                    const colPrev = document.getElementById('total-preventiva')
                    colPrev.append(tdPrev)

                    const tdConPrev = document.createElement('td')
                    tdConPrev.innerHTML = dataPrev.filter(statusConcluida).length
                    const colConPrev = document.getElementById('total-preventiva')
                    colConPrev.append(tdConPrev)

                    const tdIniPrev = document.createElement('td')
                    tdIniPrev.innerHTML = dataPrev.filter(statusIniciada).length
                    const colIniPrev = document.getElementById('total-preventiva')
                    colIniPrev.append(tdIniPrev)

                    const tdNinPrev = document.createElement('td')
                    tdNinPrev.innerHTML = dataPrev.filter(statusNaoIniciada).length
                    const colNinPrev = document.getElementById('total-preventiva')
                    colNinPrev.append(tdNinPrev)

                    const tdSumPrev = document.createElement('td')
                    tdSumPrev.innerHTML = dataPrev.filter(conIniNin).length
                    const colSumPrev = document.getElementById('total-preventiva')
                    colSumPrev.append(tdSumPrev)



                    // ============= btn Download

                    const btnProducao = document.createElement('button')
                    btnProducao.id = 'btnDonwload'
                    btnProducao.innerHTML = 'BAIXAR IMAGEM'
                    const btnProducao2 = document.getElementById('div-download')
                    btnProducao2.append(btnProducao)


                    let btnGenerator = document.querySelector('#btnDonwload')
                    let btnDownload = document.querySelector('.download')

                    btnGenerator.addEventListener('click', () => {
                         html2canvas(document.querySelector("#canvasDown")).then(canvas => {
                              document.body.appendChild(canvas)
                              btnDownload.href = canvas.toDataURL('image/png');
                              btnDownload.download = 'producao';
                              btnDownload.click();
                         });
                    })



               });



          }
     }

})





