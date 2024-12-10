import comtypes.client
import sys
#-----------------------------------------




class KTP_EtabModel:
	def __init__(self):
		self.EtabsModel, self.Etabs_file_full_name, self.Etabs_file_name = self.KTP_connect_etab()
		self.name = self.Etabs_file_name
		self.author = "Nguyen Van Chieu and Jame Hoang"
		self.model = self.EtabsModel

	def KTP_connect_etab(self):
		# print("---- Bắt Đầu Hàm connect_etab ----")
		try:
			#get the active ETABS object
			EtabsObject=comtypes.client.GetActiveObject("CSI.ETABS.API.ETABSObject")
			self.EtabsModel=EtabsObject.SapModel
			self.Etabs_file_full_name = self.EtabsModel.GetModelFilename()
			self.Etabs_file_name = self.Etabs_file_full_name.split("\\")[-1]
			kN_m_C = 6
			ret = self.EtabsModel.SetPresentUnits(kN_m_C)
			return self.EtabsModel, self.Etabs_file_full_name, self.Etabs_file_name
		except (OSError,comtypes.COMError):
			print("No running instance etabs of the program found or failed to attach.")
			self.EtabsModel = 1
			self.Etabs_file_full_name = 1
			self.Etabs_file_name = 1
			return self.EtabsModel, self.Etabs_file_full_name, self.Etabs_file_name
	def KTP_create_frame(self,P1, P2, Section):
		frame = self.model.FrameObj.AddByCoord(P1[0],P1[1],P1[2],P2[0],P2[1],P2[2], "", Section)
		print("Create beam successfully!!!!")
		return frame


