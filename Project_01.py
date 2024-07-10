
import win32com.client as wincom

if __name__ == '__main__':
    print("Welcome to Robot speake")
    while True:
        x = input("Enter what you want me to speak: ")
        if x == "Q":
            print("OK!!.. let's Bye ")
            break
        
        speak = wincom.Dispatch("SAPI.spVoice")
        speak.Speak(x)        
        