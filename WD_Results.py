import pandas as pd
import glob
import shutil

 

excel_data_path = "C:/Users/AST-TITAN/Documents/Matrox/Data/New_Data_5_18_Labeled/*.xlsx"

source_folder = r"C:/Users/AST-TITAN/Documents/Matrox/Data/New_Data_5_18_Labeled/Labeler/"

destination_folder_M0 = r"C:/Users/AST-TITAN/Documents/Matrox/Data/New_Data_5_18_Labeled/Python Sliding Window Results/M0/"
destination_folder_A5 = r"C:/Users/AST-TITAN/Documents/Matrox/Data/New_Data_5_18_Labeled/Python Sliding Window Results/A5/"
destination_folder_A2 = r"C:/Users/AST-TITAN/Documents/Matrox/Data/New_Data_5_18_Labeled/Python Sliding Window Results/A2/"
destination_folder_A8 = r"C:/Users/AST-TITAN/Documents/Matrox/Data/New_Data_5_18_Labeled/Python Sliding Window Results/A8/"
destination_folder_A11 = r"C:/Users/AST-TITAN/Documents/Matrox/Data/New_Data_5_18_Labeled/Python Sliding Window Results/A11/"
destination_folder_A10 = r"C:/Users/AST-TITAN/Documents/Matrox/Data/New_Data_5_18_Labeled/Python Sliding Window Results/A10/"
destination_folder_A3 = r"C:/Users/AST-TITAN/Documents/Matrox/Data/New_Data_5_18_Labeled/Python Sliding Window Results/A3/"
#destination_folder_good = r"C:/Users/AST-TITAN/Documents/Matrox/Data/New_Data_5_18_Labeled/Python Sliding Window Results/Good/"
#destination_folder_Defect = r"C:/Users/AST-TITAN/Documents/Matrox/Data/New_Data_5_18_Labeled/Python Sliding Window Results/Defect/" 

 

for f in glob.glob(excel_data_path):
    df = pd.read_excel(f)
    
    for i in range(0, len(df), 7):
        slc = df.iloc[i : i + 7]
        pred = slc[slc["Class scores"] == slc["Class scores"].max()]
        # class_score = slc['Class scores']
        # pred = class_score.max()
        # print(pred)

        M0 = pred.loc[(pred['Class Index'] == '1_M0 0')]
        A5 = pred.loc[(pred['Class Index'] == '2_A5 1')]
        A2 = pred.loc[(pred['Class Index'] == '3_A2 2')]
        A8 = pred.loc[(pred['Class Index'] == '4_A8 3')]
        A11 = pred.loc[(pred['Class Index'] == '5_A11 4')]
        A10 = pred.loc[(pred['Class Index'] == '6_A10 5')]
        A3 = pred.loc[(pred['Class Index'] == '7_A3 6')]
        # Good = pred.loc[(pred['Class Index'] == 'Good 7')]      
        # Defect = pred.loc[(pred['Class Index'] == '1_Label_Defect_Tiles 0')]
      
        M0_defect_list = M0.iloc[:,0]
        A5_defect_list = A5.iloc[:,0]
        A2_defect_list = A2.iloc[:,0]
        A8_defect_list = A8.iloc[:,0]
        A11_defect_list = A11.iloc[:,0]
        A10_defect_list = A10.iloc[:,0]
        A3_defect_list = A3.iloc[:,0]
        # Good_defect_list = Good.iloc[:,0]
        # Defect_defect_list = Defect.iloc[:,0] 

        M0_string = M0_defect_list.to_string(index=False)
        M0_tile_image = M0_string.splitlines(True) 

        A5_string = A5_defect_list.to_string(index=False)
        A5_tile_image = A5_string.splitlines(True) 

        A2_string = A2_defect_list.to_string(index=False)
        A2_tile_image = A2_string.splitlines(True) 

        A8_string = A8_defect_list.to_string(index=False)
        A8_tile_image = A8_string.splitlines(True) 

        A11_string = A11_defect_list.to_string(index=False)
        A11_tile_image = A11_string.splitlines(True) 

        A10_string = A10_defect_list.to_string(index=False)
        A10_tile_image = A10_string.splitlines(True) 

        A3_string = A3_defect_list.to_string(index=False)
        A3_tile_image = A3_string.splitlines(True) 

        # Good_string = Good_defect_list.to_string(index=False)
        # Good_tile_image = Good_string.splitlines(True)

        # Defect_string = Defect_defect_list.to_string(index=False)
        # Defect_tile_image = Defect_string.splitlines(True)
        
        for tile in M0_tile_image:
            name = tile.strip()
            source = source_folder + name
            # source = source[:len(source)-1]
            destination = destination_folder_M0 + name
            # destination = destination[:len(destination)-1]
            try:
                shutil.move(source, destination)
                print('Moved:', name)
            except:
                pass

        for tile in A5_tile_image:
            name = tile.strip()
            source = source_folder + name
            # source = source[:len(source)-1]
            destination = destination_folder_A5 + name
            # destination = destination[:len(destination)-1]
            try:
                shutil.move(source, destination)
                print('Moved:', name)
            except:
                pass

        for tile in A2_tile_image:
            name = tile.strip()
            source = source_folder + name
            # source = source[:len(source)-1]
            destination = destination_folder_A2 + name
            # destination = destination[:len(destination)-1]
            try:
                shutil.move(source, destination)
                print('Moved:', name)
            except:
                pass 

        for tile in A8_tile_image:
            name = tile.strip()
            source = source_folder + name
            # source = source[:len(source)-1]
            destination = destination_folder_A8 + name
            # destination = destination[:len(destination)-1]
            try:
                shutil.move(source, destination)
                print('Moved:', name)
            except:
                pass 

        for tile in A11_tile_image:
            name = tile.strip()
            source = source_folder + name
            # source = source[:len(source)-1]
            destination = destination_folder_A11 + name
            # destination = destination[:len(destination)-1]
            try:
                shutil.move(source, destination)
                print('Moved:', name)
            except:
                pass
  
        for tile in A10_tile_image:
            name = tile.strip()
            source = source_folder + name
            # source = source[:len(source)-1]
            destination = destination_folder_A10 + name
            # destination = destination[:len(destination)-1]
            try:
                shutil.move(source, destination)
                print('Moved:', name)
            except:
                pass

        for tile in A3_tile_image:
            name = tile.strip()
            source = source_folder + name
            # source = source[:len(source)-1]
            destination = destination_folder_A3 + name
            # destination = destination[:len(destination)-1]
            try:
                shutil.move(source, destination)
                print('Moved:', name)
            except:
                pass
            
        # for tile in Good_tile_image:
        #     name = tile.strip()
        #     source = source_folder + name
        #     # source = source[:len(source)-1]
        #     destination = destination_folder_good + name
        #     # destination = destination[:len(destination)-1]
        #     try:
        #         shutil.move(source, destination)
        #         print('Moved:', name)
        #     except:
        #         pass
            
        # for tile in Defect_tile_image:
        #     name = tile.strip()
        #     source = source_folder + name
        #     # source = source[:len(source)-1]
        #     destination = destination_folder_Defect + name
        #     # destination = destination[:len(destination)-1]
        #     try:
        #         shutil.move(source, destination)
        #         print('Moved:', name)
        #     except:
        #         pass