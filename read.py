def createFile(filename):
    # Open file in write mode
    with open("path.txt", "w") as file:
        file.write(filename)
    # return path

def readFile(filename):

    # Open file in read mode
    with open(filename, "r") as file:
        path = file.readlines()
        # print(path[0])
    return path[0]

