import pandas as pd
virDict = {
    'rhinovirus': [],
    'enterovirus': [],
    'influenza':[],
    'respiratory syncytial virus': [],
    'coronavirus': [],
    'adenovirus': [],
    'parainfluenza': [],
    'bocavirus': [],
    'metapneumovirus': [],
    'polyomavirus': [],
    'parechovirus': [],
    'mumps': [],
    'measles': [],
    'middle east respiratory syndrome coronavirus': [],
    'parvovirus':[]
}
animals = ['baboon','dolphin','bovine','chimpanzee','whale','chimpanzee','avian','bulbul','canine','chinese bamboo rat',
           'ferret','common-moorhen','duck','camel','equine','eurasian oystercatcher','feline','fowl','goose','seal',
           'giraffe','white-eye','yellow-bellied weasel','shorebird','turkey','sparrow','swine','rhinolophus','wigeon','rat',
           'mouse','gull','cat','rabbit','quail','puffinosis','porcine','pigeon','pheasant','night-heron','murine','munia',
           'mink','mallard','magpie-robin','raccoon dog','pintail','myotis davidii','myotis daubentonii','parrot','rock sandpiper',
           'mystacina','spotted hyena'
           ]

def lowercase(ToList):
    '''
    A function to transform data into lowercase
    '''
    ToList_copy = ToList[:]
    ToList_low = []
    for name in ToList_copy:
        split_temp = name.split(' ')
        Str_temp = ' '.join(split_temp[1:])
        ToList_low.append(Str_temp.lower())
    return ToList_low

def flu_separator(animals):
    '''
    The function to add new animal string ino animals list from other species flu
    Args:
        animals: a exist list of str of animal to filter the dictionary
    return:
        animals: updated list
    '''
    animal_df = pd.read_excel(r'C:\Users\s9012\iCloudDrive\Coding\1. Python\Roche probe\Respiratory virus\Other species.xlsx')
    animalList = list(animal_df.iloc[:,0])
    for i in range(len(animalList)):
        animal_des = animalList[i].split('/')
        animal_low = animal_des[1].lower()
        if animal_low not in animals:
            animals.append(animal_low)
            
def selection(animals,ToList_low,ToList,virDict):
    '''
    The function is to select the human respiratory virus into virDict
    Args:
        animals: a exist list of str of animal to filter the dictionary
        ToList_low: a list of lowercase data
        ToList: the original list of data
        virDict: empty dictionary for classification
    '''
    
    for i in range(len(ToList_low)):
        for name in virDict:
            # specifically select with human title
            if name == 'adenovirus' or name == 'parainfluenza' or name == 'parvovirus' or name == 'parechovirus'\
                    or name == 'polyomavirus' or name == 'metapneumovirus' or name == 'bocavirus' or name == 'respiratory syncytial virus':
                if name in ToList_low[i] and 'human' in ToList_low[i]:
                    virDict[name].append(ToList[i])
                    
            # select influenza and exclude parainfluenza and other species
            elif name == 'influenza' and name in ToList_low[i]:
                if 'parainfluenza' not in ToList_low[i]:
                    if any(animal in ToList_low[i] for animal in animals):
                        pass
                    else:
                        virDict[name].append(ToList[i])
                        
            # select enterovirus and exclude other species
            elif name == 'enterovirus' and name in ToList_low[i]:
                if any(animal in ToList_low[i] for animal in animals):
                    pass
                else:
                    virDict[name].append(ToList[i])

            # select coronavirus and name both contained 'bat' and 'sars', and exclude other species
            elif name == 'coronavirus':
                if name in ToList_low[i] or 'sars' in ToList_low[i]:
                    if any(animal in ToList_low[i] for animal in animals):
                        pass
                    else:
                        if 'bat' in ToList_low[i]:
                            if 'sars' in ToList_low[i]:
                                virDict[name].append(ToList[i])
                        else:
                            virDict[name].append(ToList[i])
            elif name in ToList_low[i]:
                virDict[name].append(ToList[i])
            
def OutputExcel (dt, file_name):
    '''
    The function transform dictionary into dataframe and output an excel with key as sheet name
    Args:
        dt: a selected dictionary
        file_name: a name for output excel
    '''
    Result_PATH = 'C:/Users/s9012/iCloudDrive/Coding/1. Python/Roche probe/Respiratory virus/' + file_name + '.xlsx'
    writer = pd.ExcelWriter(Result_PATH, engine='xlsxwriter')
    for key in dt:
        if key == 'middle east respiratory syndrome coronavirus':
            df = pd.DataFrame(dt[key],columns=['MERS'])
            df.to_excel(writer,sheet_name = str('MERS'), index=False)
        else:
            df = pd.DataFrame(dt[key], columns=[str.capitalize(key)])
            df.to_excel(writer, sheet_name=str(str.capitalize(key)), index=False)
    writer.save()

if __name__ == '__main__':
    virDF = pd.read_excel(r'C:\Users\s9012\iCloudDrive\Coding\1. Python\Roche probe\Respiratory virus\Database.xlsx',
                          sheet_name = 'vircap database')
    ToList = sorted(list(virDF.iloc[:,0]),key = lambda k:k[24:]) #sort the data by the viral name
    ToList_low = lowercase(ToList)
    flu_separator(animals) #load lots of animal name into the animals list
    selection(animals,ToList_low, ToList, virDict)
    OutputExcel(virDict,'Selected virus')