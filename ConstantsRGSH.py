
POSITION_ID =  (307499772460.0, #bTYPE ID - 0
                    4.0, #PART NUMBER ID - 1
                    5.0) #DESCRIPTION

bTYPE_NAMES = (('ASSY'),
              ('DETAIL'),
              ('STANDART'),
              ('PURCHASED'))

SOURCE_PROPERTIES_ID = (
                    (276022879831.0, 'OD'), #OD - 0
                    (296044803483.0,'WT'), #WT - 1
                    (306102640299.0, 'L_TUBE'), #L_tube - 2
                    (226778862707.0, 'T'), #tickness - 3
                    (235833998283.0, 'L_PROFILE'), #L_profile - 4
                    (334229340093.0, 'WIDTH'), #Width - 5
                    (289516577711.0, 'PROFILE_NAME'), #Name of profile - 6
                    (290108629069.0, 'PIPE_MAT'),  #Material name of pipe - 7
                    (8.0, 'MASS'), #MASS - 8
                    (307499772460.0, 'bTYPE') # bTYPE - 9
                    )

B_SPEC =        ('RGS.5.500.TUBE',
                'RGS.XXX CV1', #PROFILE CV1
                'RGS.XXX CV2') #PROFILE CV2

mRGS_PIPING = (['API 5L X65', 73.0, 'RGS 5.402.2.30'], #M_tube, mOD, mRGS #MEG
                   ['25Cr Duplex SS', 508.0, 'RGS 5.434.4.1'],#M_tube, mOD, mRGS #20" PROD
                   ['25Cr Duplex SS', 406.4, 'RGS 5.434.4.2'],#M_tube, mOD, mRGS #16" PROD
                   ['25Cr Duplex SS', 219.1, 'RGS 5.434.3.1'],#M_tube, mOD, mRGS #8" PROD
                   ['API 5L X65', 33.4, 'RGS 5.402.2.30'],#M_tube, mOD, mRGS #MEG
                   ['316L', 19.05, 'RGS 5.431.3.5'],#M_tube, mOD, mRGS
                   ['25Cr Duplex SS', 19.05, 'RGS 5.434.6.1'],#M_tube, mOD, mRGS
                   ['316L', 48.3, 'RGS 5.431.3.5'],#M_tube, mOD, mRGS
                )

mNPS =          ([9.50, '0.375'], #OD, NPS
                [12.7,'0.5'],
                [19.05,'0.75'],
                [33.4,'1'],
                [48.3,'1.5'],
                [73.0,'2.5'],
                [219.1,'8'],
                [406.4,'16'],
                [508.0, '20'],
                [823.0,'32'])

mSCH =          ([12.7,'10'],  #OD, SCH
                [19.1,'20'],
                [33.4,'30'],
                [48.3,'40'],
                [73.0,'50'],
                [219.1,'60'],
                [406.4,'70'],
                [508.0, '75'],
                [823.0,'80'])

BOM_MTO_IDs = ( (261740131642.0, 'mNPS'), #mNPS 0
                (270518399995.0, 'mDESC'), #mDESC 1
                (5.0, 'bDESC'), #bDESC 2
                (261304942691.0, 'mDIM'), #mDIM 3
                (250919968540.0, 'mMAT'), #mMAT 4
                (260572918959.0, 'bMAT'), #bMAT 5
                (388874213208.0, 'mOD'), #mOD 6
                (274824233236.0, 'mID'), #mID 7
                (360146668533.0, 'mWT'), #mWT 8
                (268876019188.0, 'mSCH'), #mSCH 9
                (266883427725.0, 'mRGS'), #mRGS 10
                (248536540390.0, 'mUNIT'), #mUnit 11
                (249942599097.0, 'mVALUE'), #mValue 12
                (250981089159.0, 'bSPEC'), #bSpec 13
                (245666335656.0, 'mST'), #mST 14
                (307499772460.0, 'bTYPE')) #bType 15

SHEET_FORMATS = [['A0', 1184, 836],
                ['A1', 836, 589],
                ['A2', 589, 415],
                ['A3', 415, 292],
                ['A4', 205, 292],]