class Headers:
    # define temp headers
    tempRMLS = 'daysDiff_RMLS'
    tempTermin = 'daysDiff_Termin'
    # define headers
    rmls = 'SollRückmeldetermin Leitstand'      # dd.mm.YYYY
    kTermin = 'Konstruktionstermin Soll'        # dd.mm.YYYY
    kem = 'Dokument'                            # clean WAR,EBS & write HKB, GEN, GEL ; SU_?, CO_?
    combinedColumns = 'Document Status'         # Combined docStatus & orderPhase
    docStatus = 'Dokumentstatus'                # 46, 47, 42
    orderPhase = 'Auftragsphase'                # E, T
    pivotTableItem1 ='Gecikme Nedeni'
    pivotTableItem2 ='Alınacak Aksiyon'
    pivotTableItem3 ='Çalışma Tipi'
