import openpyxl

from DataTransfer import DataTransfer


class Main:

    def __init__(self):
        pass

    singleIdentity = DataTransfer()
    group_ID_Identity = DataTransfer()
    group_Dif_Identity = DataTransfer()

    Gender = DataTransfer()
    Gender = DataTransfer()

    # Gender.transfer(1, "CompilationGender.xlsx", "Sheet1", "G_1.xlsx", "GENDER_01", 2,
    #                            30, 7, "CompilationGender.xlsx")
    # singleIdentity.transfer(1, "CompilationIdentity.xlsx", "Sheet1", "Exp_6_Diff_Ident_01.xlsx",
    #                                   "Exp_2_O1_Deleted.csv", 62, 21,9,"CompilationIdentity.xlsx")

    # for x in range(1,34):
    #
    #     if(x>9):
    #         singleIdentity.transfer(x, "CompilationIdentity.xlsx", "Sheet1", "Exp_6_Diff_Ident_"+str(x)+".xlsx",
    #                                  "Exp_2_O1_Deleted.csv", 2, 16,9,"CompilationIdentity.xlsx")
    #         group_ID_Identity.transfer(x, "CompilationIdentity.xlsx", "Sheet1", "Exp_6_Diff_Ident_"+str(x)+".xlsx",
    #                                    "Exp_2_O1_Deleted.csv", 62, 21,9,"CompilationIdentity.xlsx")
    #         group_Dif_Identity.transfer(x, "CompilationIdentity.xlsx", "Sheet1", "Exp_6_Diff_Ident_" + str(x) + ".xlsx",
    #                                    "Exp_2_O1_Deleted.csv", 122, 26,9,"CompilationIdentity.xlsx")
    #     else:
    #         singleIdentity.transfer(x, "CompilationIdentity.xlsx", "Sheet1", "Exp_6_Diff_Ident_0"+str(x)+".xlsx",
    #                                  "Exp_2_O1_Deleted.csv", 2, 16,9,"CompilationIdentity.xlsx")
    #         group_ID_Identity.transfer(x, "CompilationIdentity.xlsx", "Sheet1", "Exp_6_Diff_Ident_0"+str(x)+".xlsx",
    #                                    "Exp_2_O1_Deleted.csv", 62, 21,9,"CompilationIdentity.xlsx")
    #         group_Dif_Identity.transfer(x, "CompilationIdentity.xlsx", "Sheet1", "Exp_6_Diff_Ident_0" + str(x) + ".xlsx",
    #                                     "Exp_2_O1_Deleted.csv", 122, 26,9,"CompilationIdentity.xlsx")
    # for x in range(38, 39):
    #     print "Activate HACKER CODE"
    #     Gender.transfer(x, "CompilationGender.xlsx", "Sheet1", "G_99.xlsx","GENDER_01", 2,
    #                           30, 7, "CompilationGender.xlsx")
    #
    #     Gender.transfer(x, "CompilationGender.xlsx", "Sheet1", "G_99.xlsx", "GENDER_01", 62,
    #                               36, 7, "CompilationGender.xlsx")
    #
    #     Gender.transfer(x, "CompilationGender.xlsx", "Sheet1", "G_99.xlsx", "GENDER_01", 122,
    #                               42, 7, "CompilationGender.xlsx")
    #
    #     Gender.transfer(x, "CompilationGender.xlsx", "Sheet1", "G_99.xlsx", "GENDER_01", 182,
    #                               48, 7, "CompilationGender.xlsx")
    #
    #     Gender.transfer(x, "CompilationGender.xlsx", "Sheet1", "G_99.xlsx", "GENDER_01", 242,
    #                    54, 7, "CompilationGender.xlsx")

    # for x in range(1, 32):
    #     if(x == 26):
    #         continue
    #     if(x<10):
    #         Gender.transfer(x, "CompilationNoise.xlsx", "Sheet1", "Exp_3_Noise_0"+str(x)+".xlsx",
    #                         "Exp_2_O1_Deleted.csv", 2, 30, 9, "CompilationNoise.xlsx")
    #         Gender.transfer(x, "CompilationNoise.xlsx", "Sheet1", "Exp_3_Noise_0" + str(x) + ".xlsx",
    #                         "Exp_2_O1_Deleted.csv", 72, 35, 9, "CompilationNoise.xlsx")
    #     else:
    #         Gender.transfer(x, "CompilationNoise.xlsx", "Sheet1", "Exp_3_Noise_" + str(x) + ".xlsx",
    #                         "Exp_2_O1_Deleted.csv", 2, 30, 9, "CompilationNoise.xlsx")
    #         Gender.transfer(x, "CompilationNoise.xlsx", "Sheet1", "Exp_3_Noise_" + str(x) + ".xlsx",
    #                        "Exp_2_O1_Deleted.csv", 72, 35, 9, "CompilationNoise.xlsx")

    # for x in range(1, 32):
    #     if(x<10):
    #         Gender.transfer(x, "CompilationFear.xlsx", "Sheet1", "Exp_4_Fear_0"+str(x)+".xlsx",
    #                         "Exp_2_O1_Deleted.csv", 2, 39, 9, "CompilationFear.xlsx")
    #         Gender.transfer(x, "CompilationFear.xlsx", "Sheet1", "Exp_4_Fear_0" + str(x) + ".xlsx",
    #                         "Exp_2_O1_Deleted.csv", 62, 44, 9, "CompilationFear.xlsx")
    #     else:
    #         Gender.transfer(x, "CompilationFear.xlsx", "Sheet1", "Exp_4_Fear_" + str(x) + ".xlsx",
    #                         "Exp_2_O1_Deleted.csv", 2, 39, 9, "CompilationFear.xlsx")
    #         Gender.transfer(x, "CompilationFear.xlsx", "Sheet1", "Exp_4_Fear_" + str(x) + ".xlsx",
    #                        "Exp_2_O1_Deleted.csv", 62, 44, 9, "CompilationFear.xlsx")

    # for x in range(1, 32):
    #     if(x<10):
    #         Gender.transfer(x, "CompilationSize.xlsx", "Sheet1", "Exp_2_GroupSize_0"+str(x)+".xlsx",
    #                         "Exp_2_O1_Deleted.csv", 2, 39, 9, "CompilationSize.xlsx")
    #         Gender.transfer(x, "CompilationSize.xlsx", "Sheet1", "Exp_2_GroupSize_0" + str(x) + ".xlsx",
    #                         "Exp_2_O1_Deleted.csv", 62, 44, 9, "CompilationSize.xlsx")
    #     else:
    #         Gender.transfer(x, "CompilationSize.xlsx", "Sheet1", "Exp_2_GroupSize_" + str(x) + ".xlsx",
    #                         "Exp_2_O1_Deleted.csv", 2, 39, 9, "CompilationSize.xlsx")
    #         Gender.transfer(x, "CompilationSize.xlsx", "Sheet1", "Exp_2_GroupSize_" + str(x) + ".xlsx",
    #                        "Exp_2_O1_Deleted.csv", 62, 44, 9, "CompilationSize.xlsx")

    # for x in range(101, 160):
    #     if(x >101 and x < 104 or x == 118 or x == 156):
    #         continue
    #
    #     Gender.transfer(x-100, "CompilationCvS.xlsx", "Sheet1", str(x)+"_emg_crowd_ex1.xlsx", "Sheet1",
    #                         2, 28, 8, "CompilationCvS.xlsx")
    #
    #     Gender.transfer(x - 100, "CompilationCvS.xlsx", "Sheet1", str(x) + "_emg_crowd_ex1.xlsx", "Sheet1",
    #                     52, 33, 8, "CompilationCvS.xlsx")
    #
    #     Gender.transfer(x - 100, "CompilationCvS.xlsx", "Sheet1", str(x) + "_emg_crowd_ex1.xlsx", "Sheet1",
    #                     102, 38, 8, "CompilationCvS.xlsx")
    #
    #     Gender.transfer(x - 100, "CompilationCvS.xlsx", "Sheet1", str(x) + "_emg_crowd_ex1.xlsx", "Sheet1",
    #                     152, 43, 8, "CompilationCvS.xlsx")
    #
    # for x in range(161, 187):
    #
    #     Gender.transfer(x-99, "CompilationCvS.xlsx", "Sheet1", str(x)+"_emg_crowd_ex1.xlsx", "Sheet1",
    #                         2, 28, 8, "CompilationCvS.xlsx")
    #
    #     Gender.transfer(x-99, "CompilationCvS.xlsx", "Sheet1", str(x) + "_emg_crowd_ex1.xlsx", "Sheet1",
    #                     52, 33, 8, "CompilationCvS.xlsx")
    #
    #     Gender.transfer(x-99, "CompilationCvS.xlsx", "Sheet1", str(x) + "_emg_crowd_ex1.xlsx", "Sheet1",
    #                     102, 38, 8, "CompilationCvS.xlsx")
    #
    #     Gender.transfer(x-99, "CompilationCvS.xlsx", "Sheet1", str(x) + "_emg_crowd_ex1.xlsx", "Sheet1",
    #                     152, 43, 8, "CompilationCvS.xlsx")

    Gender.transfer(60, "CompilationCvS.xlsx", "Sheet1", "160a_emg_crowd_ex1.xlsx", "Sheet1",
                    2, 28, 8, "CompilationCvS.xlsx")
    Gender.transfer(60, "CompilationCvS.xlsx", "Sheet1", "160a_emg_crowd_ex1.xlsx", "Sheet1",
                    52, 33, 8, "CompilationCvS.xlsx")
    Gender.transfer(60, "CompilationCvS.xlsx", "Sheet1", "160a_emg_crowd_ex1.xlsx", "Sheet1",
                    102, 38, 8, "CompilationCvS.xlsx")
    Gender.transfer(60, "CompilationCvS.xlsx", "Sheet1", "160a_emg_crowd_ex1.xlsx", "Sheet1",
                    152, 43, 8, "CompilationCvS.xlsx")

    Gender.transfer(61, "CompilationCvS.xlsx", "Sheet1", "160b_emg_crowd_ex1.xlsx", "Sheet1",
                    2, 28, 8, "CompilationCvS.xlsx")
    Gender.transfer(61, "CompilationCvS.xlsx", "Sheet1", "160b_emg_crowd_ex1.xlsx", "Sheet1",
                    52, 33, 8, "CompilationCvS.xlsx")
    Gender.transfer(61, "CompilationCvS.xlsx", "Sheet1", "160b_emg_crowd_ex1.xlsx", "Sheet1",
                    102, 38, 8, "CompilationCvS.xlsx")
    Gender.transfer(61, "CompilationCvS.xlsx", "Sheet1", "160b_emg_crowd_ex1.xlsx", "Sheet1",
                    152, 43, 8, "CompilationCvS.xlsx")





