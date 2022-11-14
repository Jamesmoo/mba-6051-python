# import the file where all the analysis will be done
import statistical_analisys as sa

def print_hi(name):
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    print_hi('PyCharm')
    # call the main file that runs the statistical analysis for the project
    # https://stackoverflow.com/questions/7974849/how-can-i-make-one-python-file-run-another
    sa.start()
