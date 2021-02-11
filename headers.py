class Headers:

    # define temp headers
    tempRMLS = 'daysDiff_RMLS'
    tempTermin = 'daysDiff_Termin'
    combinedColumns = 'Document Status'  # Combined docStatus & orderPhase

    # define headers
    rmls = 'SollRückmeldetermin Leitstand'  # dd.mm.YYYY
    kTermin = 'Konstruktionstermin Soll'  # dd.mm.YYYY
    docType = 'Dokument'  # clean WAR,EBS & write HKB, GEN, GEL ; SU_?, CO_?
    bbNummer = 'BB-Nummer'
    docStatus = 'Dokumentstatus'  # 46, 47, 42, FG
    orderPhase = 'Auftragsphase'  # E, T
    pivotTableItem1 = 'Gecikme Nedeni'
    pivotTableItem2 = 'Alınacak Aksiyon'
    pivotTableItem3 = 'Çalışma Tipi'
