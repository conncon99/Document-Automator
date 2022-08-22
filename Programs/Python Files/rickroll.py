import webbrowser
import xlwings as xw

if __name__ == "__main__":
    xw.Book("automation.xlsm").set_mock_caller()
    webbrowser.open('https://www.youtube.com/watch?v=dQw4w9WgXcQ', new=2)
    print("Done!")
