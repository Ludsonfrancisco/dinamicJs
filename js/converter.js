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
                              item['Habilidades de Trabalho'] === 'Reparo TV FB Alto Valor(1/100)')

                              return item
                    }

                    //===================== DADOS CIDADES                   
                    let dataArc = rowObject.filter(item => item.Cidade === 'ARACRUZ')
                    let dataCim = rowObject.filter(item => item.Cidade === 'CACHOEIRO DE ITAPEMIRIM')
                    let dataCca = rowObject.filter(item => item.Cidade === 'CARIACICA')
                    let dataCna = rowObject.filter(item => item.Cidade === 'COLATINA')
                    let dataGri = rowObject.filter(item => item.Cidade === 'GUARAPARI')
                    let dataLns = rowObject.filter(item => item.Cidade === 'LINHARES')
                    let dataSmt = rowObject.filter(item => item.Cidade === 'SAO MATEUS')
                    let dataSea = rowObject.filter(item => item.Cidade === 'SERRA')
                    let dataVva = rowObject.filter(item => item.Cidade === 'VILA VELHA')
                    let dataVta = rowObject.filter(item => item.Cidade === 'VITORIA')



                    //  DADOS CIDADE METALICO =========================
                    const dataMetalico = rowObject.filter(metalico)
                    const metalicoCca = dataCca.filter(metalico)
                    const metalicoCna = dataCna.filter(metalico)
                    const metalicoLns = dataLns.filter(metalico)
                    const metalicoSea = dataSea.filter(metalico)
                    const metalicoVva = dataVva.filter(metalico)
                    const metalicoVta = dataVta.filter(metalico)

                    console.log(dataMetalico.length)
                    // ============= DADOS POR CIDADE GPON 

                    const dataGpon = rowObject.filter(gpon)
                    const gponArc = dataArc.filter(gpon)
                    const gponCim = dataCim.filter(gpon)
                    const gponCca = dataCca.filter(gpon)
                    const gponCna = dataCna.filter(gpon)
                    const gponGri = dataGri.filter(gpon)
                    const gponLns = dataLns.filter(gpon)
                    const gponSmt = dataSmt.filter(gpon)
                    const gponSea = dataSea.filter(gpon)
                    const gponVva = dataVva.filter(gpon)
                    const gponVta = dataVta.filter(gpon)

                    // ====CREATE CABEÇALHO TABELA PRODUÇAO GPON ====

                    const titleGpon = document.createElement('span')
                    titleGpon.innerHTML = 'GPON'
                    const tGpon = document.getElementById("title-gpon")
                    tGpon.append(titleGpon)

                    const tdCidade = document.createElement('td')
                    tdCidade.innerHTML = 'CIDADE'
                    const tdConcluida = document.createElement('td')
                    tdConcluida.innerHTML = 'CONCLUIDA'
                    const tdIniciada = document.createElement('td')
                    tdIniciada.innerHTML = 'INICIADA'
                    const tdNaoiniciada = document.createElement('td')
                    tdNaoiniciada.innerHTML = 'NÃO INICIADA'
                    const total = document.createElement('td')
                    total.innerHTML = 'TOTAL'

                    const tabela = document.getElementById('cabecalho')
                    tabela.append(tdCidade)
                    tabela.append(tdConcluida)
                    tabela.append(tdIniciada)
                    tabela.append(tdNaoiniciada)
                    tabela.append(total)

                    // ======== CREATE CIDADE GPON ====


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
                    tdCimGpon.innerHTML = 'CACHOEIRO DE ITAPEMIRIM'
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
                    const colCnaGpon = document.getElementById('cna-metalico')
                    colCnaGpon.append(tdCnaGpon)

                    const tdConCnaGpon = document.createElement('td')
                    tdConCnaGpon.innerHTML = metalicoCna.filter(statusConcluida).length
                    const colConCnaGpon = document.getElementById('cna-metalico')
                    colConCnaGpon.append(tdConCnaGpon)

                    const tdIniCnaGpon = document.createElement('td')
                    tdIniCnaGpon.innerHTML = metalicoCna.filter(statusIniciada).length
                    const colIniCnaGpon = document.getElementById('cna-metalico')
                    colIniCnaGpon.append(tdIniCnaGpon)

                    const tdNinCnaGpon = document.createElement('td')
                    tdNinCnaGpon.innerHTML = metalicoCna.filter(statusNaoIniciada).length
                    const colNinCnaMetalico = document.getElementById('cna-metalico')
                    colNinCnaMetalico.append(tdNinCnaGpon)

                    const tdTotalCnaMetalico = document.createElement('td')
                    tdTotalCnaMetalico.innerHTML = metalicoCna.filter(conIniNin).length
                    const colTotalCnaMetalico = document.getElementById('cna-metalico')
                    colTotalCnaMetalico.append(tdTotalCnaMetalico)


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


                    // TOTAL
                    const tdGpon = document.createElement('td')
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


                    // ============ CREATE CIDADADE METALICO                 


                    const titleMetalico = document.createElement('span')
                    titleMetalico.innerHTML = 'METALICO'
                    const tMetalico = document.getElementById("title-metalico")
                    tMetalico.append(titleMetalico)

                    const tdCidadeMetalico = document.createElement('td')
                    tdCidadeMetalico.innerHTML = 'CIDADE'
                    const tdConcluidaMetalico = document.createElement('td')
                    tdConcluidaMetalico.innerHTML = 'CONCLUIDA'
                    const tdIniciadaMetalico = document.createElement('td')
                    tdIniciadaMetalico.innerHTML = 'INICIADA'
                    const tdNaoiniciadaMetalico = document.createElement('td')
                    tdNaoiniciadaMetalico.innerHTML = 'NÃO INICIADA'
                    const totalMetalico = document.createElement('td')
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
                    const colCnaMetalico = document.getElementById('cna')
                    colCnaMetalico.append(tdCnaMetalico)

                    const tdConCnaMetalico = document.createElement('td')
                    tdConCnaMetalico.innerHTML = gponCna.filter(statusConcluida).length
                    const colConCnaMetalico = document.getElementById('cna')
                    colConCnaMetalico.append(tdConCnaMetalico)

                    const tdIniCnaMetalico = document.createElement('td')
                    tdIniCnaMetalico.innerHTML = gponCna.filter(statusIniciada).length
                    const colIniCnaMetalico = document.getElementById('cna')
                    colIniCnaMetalico.append(tdIniCnaMetalico)

                    const tdNinCnaMetalico = document.createElement('td')
                    tdNinCnaMetalico.innerHTML = gponCna.filter(statusNaoIniciada).length
                    const colNinCnaGpon = document.getElementById('cna')
                    colNinCnaGpon.append(tdNinCnaMetalico)

                    const tdTotalCnaGpon = document.createElement('td')
                    tdTotalCnaGpon.innerHTML = gponCna.filter(conIniNin).length
                    const colTotalCnaGpon = document.getElementById('cna')
                    colTotalCnaGpon.append(tdTotalCnaGpon)


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

                    const btnProducao = document.createElement('button')
                    btnProducao.id = 'btnDonwload'
                    btnProducao.innerHTML = 'DOWNLOAD'
                    const btnProducao2 = document.getElementById('canvasDown')
                    btnProducao2.append(btnProducao)
               
          
                    let btnGenerator = document.querySelector('#canvasDown')
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





