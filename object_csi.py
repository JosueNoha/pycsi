import comtypes.client
import pandas as pd
import numpy as np

class EtabsModel:
    units_dict = {
            'kN_mm' : 5,
            'kN_m' : 6,
            'kgf_mm' : 7,
            'kgf_m' : 8,
            'N_mm' : 9,
            'N_m' : 10,
            'tonf_mm' : 11,
            'tonf_m' : 12,
            'kn_cm' : 13,
            'kgf_cm' : 14,
            'N_cm' : 15,
            'tonf_cm' : 16,
    }

    def __init__(self):
        #create API helper object
        helper = comtypes.client.CreateObject(f'ETABSv1.Helper')
        helper = helper.QueryInterface(comtypes.gen.ETABSv1.cHelper)
        self.EtabsObject = helper.GetObject(f"CSI.ETABS.API.ETABSObject")
        self.SapModel = self.EtabsObject.SapModel
        
    def set_units(self,unit):
        SapModel = self.SapModel
        SapModel.SetPresentUnits(self.units_dict[unit])

    def close_etabs(self):
        SapModel,EtabsObject = self.SapModel,self.EtabsObject
        SapModel.SetModelIsLocked(False)
        EtabsObject.ApplicationExit(True)
        self.SapModel = None
        self.EtabsObject = None

    def lock_model(self):
        SapModel = self.SapModel
        SapModel.SetModelIsLocked(True)
        
    def unlock_model(self):
        SapModel = self.SapModel
        SapModel.SetModelIsLocked(False)

    def set_envelopes_for_dysplay(self,set_envelopes=True):
        SapModel = self.SapModel
        IsUserBaseReactionLocation=False
        UserBaseReactionX=0
        UserBaseReactionY=0
        UserBaseReactionZ=0
        IsAllModes=True
        StartMode=0
        EndMode=0
        IsAllBucklingModes=True
        StartBucklingMode=0
        EndBucklingMode=0
        MultistepStatic=1 if set_envelopes else 2
        NonlinearStatic=1 if set_envelopes else 2
        ModalHistory=1
        DirectHistory=1
        Combo=2
        SapModel.DataBaseTables.SetOutputOptionsForDisplay(IsUserBaseReactionLocation,UserBaseReactionX,
                                                                UserBaseReactionY,UserBaseReactionZ,IsAllModes,
                                                                StartMode,EndMode,IsAllBucklingModes,StartBucklingMode,
                                                                EndBucklingMode,MultistepStatic,NonlinearStatic,
                                                                ModalHistory,DirectHistory,Combo)
        
    def get_table(self,table_name,set_envelopes=True):
        SapModel = self.SapModel
        self.set_envelopes_for_dysplay(set_envelopes)
        data = self.SapModel.DatabaseTables.GetTableForDisplayArray(table_name,FieldKeyList='',GroupName='')
        
        if not data[2][0]:
            SapModel.Analyze.RunAnalysis()
            data = SapModel.DatabaseTables.GetTableForDisplayArray(table_name,FieldKeyList='',GroupName='')
            
        columns = data[2]
        data = [i if i else '' for i in data[4]] #reemplazando valores None por ''
        #reshape data
        data = pd.DataFrame(data)
        data = data.values.reshape(int(len(data)/len(columns)),len(columns))
        table = pd.DataFrame(data, columns=columns)
        return table
    
    def get_beam_forces(self,env_name="ENVOLVENTE",units='N_m'):
        try:
            SapModel = self.SapModel
            self.set_units(units)
            SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()
            SapModel.Results.Setup.SetComboSelectedForOutput(env_name)
            data = pd.DataFrame(SapModel.Results.FrameForce('',3)[1:-1]).T
            Vmax = data[data[5]=='Max'][8]
            Vmin = data[data[5]=='Min'][8]
            Mmin = data[data[5]=='Min'][12]
            Mmax = data[data[5]=='Max'][12]
            Est = data[data[5]=='Max'][1]
            datos = pd.DataFrame(np.array([Est,Vmax,Vmin,Mmax,Mmin]).T,columns=['Est.','Vmax','Vmin','Mmax','Mmin'])
            return datos
        except:
            print('Seleccione una viga')
            
    def set_concrete(self,fc=21,units='N_mm'):
        SapModel = self.SapModel
        MATERIAL_CONCRETE = 2
        N_CONCRETO = "Conc f'c_" + str(fc) + ' '+ units
        SapModel.PropMaterial.SetMaterial(N_CONCRETO, MATERIAL_CONCRETE)
        self.set_units('N_mm')
        SapModel.PropMaterial.SetMPIsotropic(N_CONCRETO, 47000*fc**0.5 , 0.2, 0.0000055)
        self.set_units(units)
        IsLightweight = False
        FcsFactor = 1
        SSType = 2 #Mander
        SSHysType = 4 #Concrete
        StrainAtFc = 0.002219
        StrainUltimate = 0.005
        SapModel.PropMaterial.SetOConcrete(N_CONCRETO,fc,IsLightweight,FcsFactor,SSType,SSHysType,StrainAtFc,StrainUltimate)

    def set_rebar(self,fy=420,fu=630, units='N_mm'):
        SapModel = self.SapModel
        self.set_units(units)
        MATERIAL_REBAR = 6
        N_REBAR = "Reb fy_" + str(fy) + " " + units
        SapModel.PropMaterial.SetMaterial(N_REBAR, MATERIAL_REBAR)
        SapModel.PropMaterial.SetMPIsotropic(N_REBAR, 2e5 , 0.2, 0.0000117)
        SSType = 1 #Simple
        SSHysType  = 1 #kinematic
        StrainAtHardening = 0.01
        StrainUltimate = 0.09
        UseCaltransSSDefaults = False
        SapModel.PropMaterial.SetORebar(N_REBAR,fy,fu,fy,fu,SSType,SSHysType,StrainAtHardening,StrainUltimate,UseCaltransSSDefaults)

    def set_beam_sections(self,b,h,r=60,fc=21,fy=420,units='N_mm'):
        SapModel = self.SapModel
        self.set_units(units)
        N_CONCRETO = "Conc f'c_" + str(fc) + ' '+ units
        N_REBAR = "Reb fy_" + str(fy) + " " + units
        N_SECCION = f'V {b} x {h} mm'
        SapModel.PropFrame.SetRectangle(N_SECCION, N_CONCRETO, h, b)
        SapModel.PropFrame.SetRebarBeam(N_SECCION,N_REBAR,N_REBAR,r,r,0.0,0.0,0.0,0.0)

    def set_column_sections(self,b,h,r=60,fc=21,fy=420,units='N_mm'):
        SapModel = self.SapModel
        self.set_units(units)
        N_CONCRETO = "Conc f'c_" + str(fc) + ' '+ units
        N_REBAR = "Reb fy_" + str(fy) + " " + units
        N_SECCION = f'C {b} x {h} mm'
        SapModel.PropFrame.SetRectangle(N_SECCION, N_CONCRETO, h, b)
        Pattern = 1 #Rectangular
        ConfineType = 1 #Ties
        NumberCBars = 0 #Circular Bars
        NumberR3Bars = 3
        NumberR2Bars = 3
        RebarSize = SapModel.PropRebar.GetNameList()[1][3]
        TieSize = SapModel.PropRebar.GetNameList()[1][1]
        TieSpacingLongit = min(b,h)/2
        Number2DirTieBars = 2
        Number3DirTieBars = 2
        ToBeDesigned = False
        SapModel.PropFrame.SetRebarColumn(N_SECCION,N_REBAR,N_REBAR,Pattern,ConfineType,r,NumberCBars,NumberR3Bars,
                                        NumberR2Bars,RebarSize,TieSize,TieSpacingLongit,Number2DirTieBars,
                                        Number3DirTieBars,ToBeDesigned)
        
    def set_shell_sections(self,h,aligerado=True,fc=21,units='N_mm'):
        SapModel = self.SapModel
        self.set_units(units)
        MatProp = "Conc f'c_" + str(fc) + ' '+ units
        if aligerado:
            name = f'aligerado e={h} mm'
            SlabType = 3
            ShellType = 3 #membrana
            Thickness = h
            SapModel.PropArea.SetSlab(name,SlabType,ShellType,MatProp,Thickness)
            SapModel.PropArea.SetSlabRibbed(name,OverallDepth=h,SlabThickness=5,StemWidthTop=10,StemWidthBot=10,
                                                RibSpacing=40,RibsParallelTo=1)

        else:
            name = f'losa e={h} mm'
            SlabType = 0
            ShellType = 1 #shell thin
            Thickness = h
            SapModel.PropArea.SetSlab(name,SlabType,ShellType,MatProp,Thickness)


    def set_wall_sections(self,e,fc=21,unit='N_mm'):
        SapModel = self.SapModel
        self.set_units(unit)
        name = f'Muro e= {e} mm'
        eWallPropType = 1
        ShellType = 1 #shell thin
        MatProp = "Conc f'c_" + str(fc) + ' '+ unit
        Thickness = e
        SapModel.PropArea.SetWall(name,eWallPropType,ShellType,MatProp,Thickness)
        

    def draw_shell(self,points,h,aligerado=True,unit='N_mm'):
        SapModel = self.SapModel
        self.set_units(unit)
        if aligerado:
            prop_name = f'aligerado e={h} mm'
        else:
            prop_name = f'losa e={h} mm'
        X = points['X']
        Y = points['Y']
        Z = points['Z']
        area_obj = SapModel.AreaObj.AddByCoord(len(X),X,Y,Z)
        area_name = area_obj[3]
        SapModel.AreaObj.SetProperty(area_name,prop_name)

    def draw_wall(self,pi,pf,e,stories='all',unit='N_mm'):
        SapModel = self.SapModel
        self.set_units(unit)
        story_data = self.get_table('Story Definitions')
        story_data.Height = story_data.Height.astype('float64')
        story_data = story_data.sort_values(by=['Story'])
        story_data['t_height'] = story_data['Height'].cumsum()
        story_data.index = story_data.Story
        if stories == 'all':
            z1 = 0
            z2 = story_data['t_height'][-1]
        elif stories[0] not in story_data['Height']:
            z1 = 0
            z2 = story_data['t_height'][stories[1]]
        else:
            z1 = story_data['t_height'][stories[0]]
            z2 = story_data['t_height'][stories[1]]
            
        X = [pi[0],pf[0],pf[0],pi[0]]
        Y = [pi[1],pf[1],pf[1],pi[1]]    
        Z = [z1,z1,z2,z2]
        prop_name = f'Muro e= {e} mm'
        area_obj = SapModel.AreaObj.AddByCoord(len(X),X,Y,Z)
        area_name = area_obj[3]
        SapModel.AreaObj.SetProperty(area_name,prop_name)

    def draw_beam(self,pi,pf,b,h,unit='N_mm'):
        SapModel = self.SapModel
        self.set_units(unit)
        Xi,Yi,Zi = pi
        Xj,Yj,Zj = pf
        propName = f'V {b} x {h} mm'
        SapModel.FrameObj.AddByCoord(Xi,Yi,Zi,Xj,Yj,Zj,PropName=propName)

    def draw_column(self,pi,b,h,stories='all',unit='N_mm'):
        SapModel = self.SapModel
        self.set_units(unit)
        X,Y = pi
        story_data = self.get_table('Story Definitions')
        story_data.Height = story_data.Height.astype('float64')
        story_data = story_data.sort_values(by=['Story'])
        story_data['t_height'] = story_data['Height'].cumsum()
        story_data.index = story_data.Story
        if stories == 'all':
            Z1 = 0
            Z2 = story_data['t_height'][-1]
        elif stories[0] not in story_data['Height']:
            Z1 = 0
            Z2 = story_data['t_height'][stories[1]]
        else:
            Z1 = story_data['t_height'][stories[0]]
            Z2 = story_data['t_height'][stories[1]]
        propName = f'C {b} x {h} mm'
        SapModel.FrameObj.AddByCoord(X,Y,Z1,X,Y,Z2,PropName=propName)


if __name__ == '__main__':
    etb = EtabsModel()
    import time
    to = time.time()
    print(etb.get_beam_forces())
    print(time.time()-to)
