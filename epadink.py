import win32com.client
import time
import sys
import os
from PIL import Image

def main():
    try:
        # Reemplaza "{TU_CLSID_AQUI}" con la CLSID específica de ePadLink
        clsid = "{84C046A7-4370-4D91-8737-87C12F4C63C5}"
        
        # Crear una instancia del objeto ActiveX
        epad = win32com.client.Dispatch(clsid)
        
        epad.HashData = "" # No obligar a Hash
        resultado = epad.ConnectedDevice # Devolver el ID del equipo epad al que se ha conectado
        print("Conectado a epadInk: ",resultado)
        if len(resultado) == 0:
            print("Conexión fallida con ePadLink.")
            sys.exit()

        print("Cuando aparezca la pantalla firme")
        # Parámetros para comenzar a firmar
        button_style =3
        sign_dig_req = 1

        resultado = epad.StartSign(button_style,sign_dig_req)

        # Parámetros para guardar la firma        
        file_name = "c:\\tmp\\epk.bmp";
        if os.path.exists(file_name):
            os.remove(file_name)
            print("Fichero ",file_name," eliminado")    

        n_width = 280
        n_height = 210
        file_type = 0;
        epad.SaveToFile(file_name,n_width,n_height,file_type)
        if os.path.exists(file_name):
            print("Firma generada en el fichero de imagen",file_name)
            imagen = Image.open(file_name)
            imagen.show()
        else:
            print("No se pudo generar la imagen de la firma")
        
        epad.ClearSign()

        epad.CloseConnection()
        print("Conexion con epadInk cerrada")


    except Exception as e:
        print("Ocurrió un error:", e)

if __name__ == "__main__":
    main()
