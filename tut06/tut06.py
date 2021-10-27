import os
import re
import shutil  
def regex_renamer(): 

    # Taking input from the user
    print("1. Breaking Bad")
    print("2. Game of Thrones")
    print("3. Lucifer")

    webseries_num = int(input("Enter the number of the web series that you wish to rename. 1/2/3: "))
    season_padding = int(input("Enter the Season Number Padding: "))
    episode_padding = int(input("Enter the Episode Number Padding: "))

    if webseries_num==1 :  
            webseries_name = 'Breaking Bad'
            #rename_Breaking_bad('Breaking Bad',season_padding,episode_padding)
            folder1 = "correct_srt"
 
            if not os.path.exists(folder1): 
                 os.makedirs(folder1)
            season_pad = int(season_padding)
            episode_pad = int(episode_padding)
            basepath = "./wrong_srt"
            destiny=os.path.join(os.getcwd(),"correct_srt",webseries_name)
            path=os.path.join(os.getcwd(),"wrong_srt",webseries_name)
            dest=os.path.join(os.getcwd(),"correct_srt")
            folder_path = os.path.join(basepath, webseries_name)
            files = os.listdir(folder_path)
            for file in files:
        
             split_file = file.split('.')
             extension = split_file[-1]
             split_file=split_file[0].split(" ")
             split_file=split_file[2]
             list_of_all_numbers_in_filename = re.findall('[0-9]+',file)
             season_num = list_of_all_numbers_in_filename[0]
             season_num = str(int(season_num))
             episode_num = list_of_all_numbers_in_filename[1]
             episode_num = str(int(episode_num))
             f_name = webseries_name + " - Season " + season_num.zfill(season_pad) + " Episode " + episode_num.zfill(episode_pad) + "." + extension
             des=os.path.join(destiny,f_name)
             if not os.path.isdir(destiny): 
                os.chdir(dest)
                
                os.mkdir("Breaking Bad")
             if not os.path.isfile(des):
                
                os.chdir(destiny)
                
                with open(f_name, 'w') as file:
                  pass
         
    elif webseries_num==2: 
          webseries_name = 'Game of Thrones'
          folder1 = "correct_srt"
 
          if not os.path.exists(folder1): 
            os.makedirs(folder1)
          season_pad = (season_padding)
          episode_pad = (episode_padding)
          basepath = "./wrong_srt"
          destiny=os.path.join(os.getcwd(),"correct_srt",webseries_name)

          path=os.path.join(os.getcwd(),"wrong_srt",webseries_name)
          dest=os.path.join(os.getcwd(),"correct_srt")
          folder_path = os.path.join(basepath, webseries_name)
          files = os.listdir(folder_path)
          for file in files:
        
            split_file = file.split('-')
            series_name = split_file[0]
            middle_split = split_file[1].split('x')
            middle_split = [ x.strip() for x in middle_split]
            season_num = middle_split[0]
            season_num = str(int(season_num))
            if len(split_file) > 3 and split_file[2].strip().isdigit():
                episode_num = str(int(middle_split[1])).zfill(episode_pad) + "-" + str(int(split_file[2].strip())).zfill(episode_pad)
                end_split = split_file[3].split('.')
                episode_name = end_split[0]
                episode_name = episode_name.strip()
                extension = file.split('.')[-1]
                f_name = series_name + "- " + "Season " + season_num.zfill(season_pad) + " Episode " + episode_num + " - " + episode_name + "." + extension
                des=os.path.join(destiny,f_name)
                if not os.path.isdir(destiny): 
                    os.chdir(dest)
                
                    os.mkdir("Game of Thrones")
                if not os.path.isfile(des):
                
                    os.chdir(destiny)
                
                    with open(f_name, 'w') as file:
                      pass
                

            else:
                episode_num = middle_split[1]
                episode_num = str(int(episode_num))
                end_split = split_file[2].split('.')
                episode_name = end_split[0]
                episode_name = episode_name.strip()
                extension = file.split('.')[-1]
    
                f_name = series_name + "- " + "Season " + season_num.zfill(season_pad) + " Episode " + episode_num.zfill(episode_pad) + " - " + episode_name + "." + extension
                des=os.path.join(destiny,f_name)
                if not os.path.isdir(destiny): 
                    os.chdir(dest)
                
                    os.mkdir("Game of Thrones")
                if not os.path.isfile(des):
                
                    os.chdir(destiny)
                
                    with open(f_name, 'w') as file:
                      pass
    
        
           
        
    

    return
regex_renamer()
