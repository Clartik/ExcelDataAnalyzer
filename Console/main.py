import functions
from colorama import init, Fore, Back, Style

# Allows color
init(convert=True)


def main():
    functions.intro()
    functions.ask()
    reclose()


# Asks if you want to restart or end the tool
def reclose():
    print("\nThank You for using the tool.")
    endAns = input(
        'If you wish to restart the tool, please type "Restart". If you wish to end the tool, please type "Close":')

    # Removes the caps sensitivity
    endAns.casefold()

    while (endAns != "restart" and endAns != "1" and endAns != "close" and endAns != "2") == True:
        print(Fore.RED + "INVAILD INPUT")
        endAns = input(
            'Please provide either "Restart" if you wish to restart the tool or "Close" if you wish to end the tool:')

    # Removes color from text
    print(Fore.RESET, end='')

    # Removes the caps sensitivity
    endAns.casefold()

    if (endAns == "restart" or endAns == "1"):
        functions.clear()
        main()
    elif (endAns == "close" or endAns == "2"):
        exit(0)


# Starts the main function at the beginning of the tool
if __name__ == "__main__":
    main()
