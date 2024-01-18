<#
.SYNOPSIS
    Config file for the Mass changing links in PPTX file PowerShell Script.
#>
@{
    # Folders : 
    Folders = @{
        # Basically on C:\temp
        ProjectFolder = 'C:\Temp'
        # The folder for PPTX, subfolder of ProjectFolder
        ProjectChangeLinks = 'ChangeLinks'
        # The working folders
        WorkingFolder = 'WORKING'
        # Destination folder after change
        DestinationFolder = 'DESTINATION'
        # Files of slides to change
        SlidesLocation = 'ppt/slides'
        SlidesRelsLocation = 'ppt/slides/_rels'
    }
    # Links to change
    AllLinks = @(
        # Repeat as needed @{}
        @{
            OldURL = 'https://service.ca-mocca.com/maia-collab/communaute_des_architectes/BigData/Gouv%20des%20donnes/Ateliers%20data%20management/'
            NewURL = 'https://casa-m.ca-mocca.com/sites/Gouvernance-Data-Groupe/AteliersDataManagement/LesAteliers/'
        }
        @{
            OldURL = 'https://service.ca-mocca.com/maia-collab/communaute_des_architectes/BigData/Gouv%20des%20donnes/Recensement%20solutions%20data%20management%20-%20ConsoGroupe.xlsx?Web=1'
            NewURL = 'https://casa-m.ca-mocca.com/sites/Gouvernance-Data-Groupe/AteliersDataManagement/Publications/Recensement_solutions_data_management.xlsx?d=w76471458b2e54c2a8986638442390e1d'
        }
        @{
            OldURL = '/20191209%20-%20atelier%20lancement/'
            NewURL = '/Lancement/'
        }
        @{
            OldURL = '/20200123%20-%20atelier%20preuve%20valeur-priorisation/'
            NewURL = '/Preuve%20valeur-priorisation/'
        }
        @{
            OldURL = '/20200305%20-%20atelier%20data%20catalogue/'
            NewURL = '/Data%20catalogue/'
        }

        @{
            OldURL = '/20200402%20-%20atelier%20data%20catalogue-dictionnaire-glossaire/'
            NewURL = '/Data%20catalogue%20-%20Dictionnaire%20-%20Glossaire/'
        }
        @{
            OldURL = '/20200515%20-%20atelier%20actu%20entites/'
            NewURL = '/Actualites%20entites/'
        }
        @{
            OldURL = '/20200625%20-%20atelier%20organisation%20-%20roles%20et%20resp'
            NewURL = '/Organisation%20-%20roles%20et%20resp/'
        }
        @{
            OldURL = '/20200710%20-%20atelier%20data%20quality/'
            NewURL = '/Data%20quality/'
        }
        @{
            OldURL = '/20200904%20-%20atelier%20gouvernance%20de%20la%20donnee%20et%20protection/'
            NewURL = '/Gouvernance%20de%20la%20donnee%20et%20protection/'
        }
        @{
            OldURL = '/20201001%20-%20atelier%20gouv%20generale%20et%20dysqualite/'
            NewURL = '/Gouvernance%20generale%20et%20dysqualite/'
        }
        @{
            OldURL = '/20201112%20-%20atelier%20lineage/'
            NewURL = '/Lineage/'
        }
        @{
            OldURL = '/20201127%20-%20atelier%20panorama%20gouv%20data%20Groupe/'
            NewURL = '/Panorama%20gouv%20data%20Groupe/'
        }
        @{
            OldURL = '/20201210%20-%20atelier%20qualite%20et%20actus%20entites/'
            NewURL = '/Qualite%20et%20actualites%20entites/'
        }
        @{
            OldURL = '/20210115%20-%20atelier%20archivage%20et%20purge/'
            NewURL = '/Archivage%20et%20purge%20-%201/'
        }
        @{
            OldURL = '/20210211%20-%20atelier%20qualite%20et%20actus/'
            NewURL = '/Qualite%20et%20actualites/'
        }
        @{
            OldURL = '/20210312%20-%20atelier%20ontologie%20-%20donner%20du%20sens%20aux%20data/'
            NewURL = '/Ontologie%20-%20donner%20du%20sens%20aux%20data/'
        }
        @{
            OldURL = '/20210611%20-%20atelier%20tdb%20KPI%20data%20management%20et%20actus%20entites/'
            NewURL = '/tdb%20KPI%20data%20management%20et%20actualites%20entites/'
        }
        @{
            OldURL = '/20210702%20-%20atelier%20glossaire%20et%20dictionnaire%20-%20actus/'
            NewURL = '/Glossaire%20et%20dictionnaire%20-%20actualites/'
        }
        @{
            OldURL = '/20210930%20-%20atelier%20qualite%20et%20gouv-orga/'
            NewURL = '/Qualite%20et%20gouv-orga/'
        }
        @{
            OldURL = '/20211118%20-%20atelier%20MDM%20-%20r%C3%A9f%C3%A9rentiels/'
            NewURL = '/MDM%20-%20r%C3%A9f%C3%A9rentiels/'
        }
        @{
            OldURL = '/20211203%20-%20atelier%20glossaire-catalogue-lineage--MDM--actus/'
            NewURL = '/Glossaire-catalogue-lineage--MDM--actus/'
        }
        @{
            OldURL = '/20220114%20-%20atelier%20qualite/'
            NewURL = '/Qualite/'
        }
        @{
            OldURL = '/20220211%20-%20catalogue-dictionnaire-lineage-restitutionsZP/'
            NewURL = '/catalogue-dictionnaire-lineage-restitutionsZP/'
        }
        @{
            OldURL = '/20220325%20-%20atelier%20maturite%20data/'
            NewURL = '/Maturite%20data/'
        }
        @{
            OldURL = '/20220513%20-%20atelier%20harmo%20data%20-%20lg%20commun/'
            NewURL = '/Harmo%20data%20-%20lg%20commun/'
        }
        @{
            OldURL = '/20220617%20-%20atelier%20harmo%20-%20lg%20commun%20-%20Groupe/'
            NewURL = '/Harmo%20-%20lg%20commun%20-%20Groupe/'
        }
        @{
            OldURL = '/20220915%20-%20atelier%20protection%20data%20-%20actus%20entites/'
            NewURL = '/Protection%20data%20-%20actus%20entites/'
        }
        @{
            OldURL = '/20221014%20-%20atelier%20REX%20gouvernance%20data%20-%20Hopex/'
            NewURL = '/REX%20gouvernance%20data%20-%20Hopex/'
        }
        @{
            OldURL = '/20221209%20-%20atelier%20archivage%20et%20purge%202/'
            NewURL = '/Archivage%20et%20purge%20-%202/'
        }
        @{
            OldURL = '/20230203%20-%20atelier%20solutions%20data%20management/'
            NewURL = '/Solutions%20data%20management%20-%201/'
        }
        @{
            OldURL = '/20230310%20-%20atelier%20solutions%20data%20management%20-%20suite/'
            NewURL = '/Solutions%20data%20management%20-%202/'
        }
        @{
            OldURL = '/20230414%20-%20atelier%20maturite%20data%202/'
            NewURL = '/Maturite%20data%202/'
        }
        @{
            OldURL = '/20230526%20-%20atelier%20data%20products/'
            NewURL = '/Data%20products/'
        }
        @{
            OldURL = '/20230623%20-%20atelier%20tour%20de%20table%20entit%C3%A9s/'
            NewURL = '/Tour%20de%20table%20entites/'
        }
        @{
            OldURL = '/20230921%20-%20atelier%20data%20gov%20Hopex%20-%20data%20marketplace/'
            NewURL = '/Data%20Gouvernance%20Hopex%20-%20Data%20Marketplace/'
        }
     )
}
# -----[EOD]------------------------------------------------------------

<#

        @{
            OldURL = ''
            NewURL = ''
        }


xxxxxxxx

https://service.ca-mocca.com/maia-collab/communaute_des_architectes/BigData/Gouv%20des%20donnes/Forms/AllItems.aspx?RootFolder=%2Fmaia%2Dcollab%2Fcommunaute%5Fdes%5Farchitectes%2FBigData%2FGouv%20des%20donnes%2FAteliers%20data%20management&FolderCTID=0x012000DBDD33F78278D2429E13D266D96CC4FE&View=%7BA461593D%2D94CE%2D40C0%2D8247%2DE36539AA8CEC%7D
https://service.ca-mocca.com/maia-collab/communaute_des_architectes/BigData/Gouv%20des%20donnes/sujets%20et%20liens%20ateliers%20data%20management.pptx?Web=1
https://service.ca-mocca.com/maia-collab/communaute_des_architectes/BigData/Partage/Forms/AllItems.aspx?RootFolder=/maia-collab/communaute_des_architectes/BigData/Partage/Data%20Management&FolderCTID=0x012000B42C3D8CA740BB46A896B9F3D2D29997&View=%7b168D9DBC-6CF0-4818-AC96-2DD1218DC802%7d


DMOG

https://service.ca-mocca.com/maia-collab/data/Cadre%20de%20Reference%20Data%20Groupe/Data%20Groupe%20-%20Cadre%20de%20R%C3%A9f%C3%A9rence%20Groupe%20-%20Politique%20Groupe%20Qualit%C3%A9%20des%20donn%C3%A9es%20-%20Version%20Pr%C3%A9liminaire.pptx?Web=1

https://service.ca-mocca.com/maia-collab/data/Cadre%20de%20Reference%20Data%20Groupe/Maturite%20Data%20-%20Mapping%20Labels.pptx?Web=1


AEG

https://service.ca-mocca.com/maia-collab/communaute_des_architectes/BigData/SI%20datacentric/Livrables/CASA-SI%20data-centric-Dossier%20d'%C3%A9tude%20SI%20Data%20Centric_V3-2.pptx?Web=1


#>