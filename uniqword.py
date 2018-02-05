import cmd
import time
import codecs
import PyRTF


class WordsFile:
    """
    Manage the file and collect and enumerate the words it contains.
    """

    all_words = []

    def __init__(self, file_path: str):
        """
        Initialise the file instance by storing a list of all words.
        :param file_path: the file path and name.
        """
        self.file_path = file_path
        self.store_all_words()

    def store_all_words(self):
        """
        Store an instance list with each word in the chosen file, eliminating every punctuation sign.
        """

        all_words = []
        with codecs.open(self.file_path) as file:
            # Flatten all lines, replace breaks with spaces and encode to UTF-8.
            contents = " ".join(file.read().splitlines())
            contents = contents.split(" ")  # Separate words.
            for word in filter(lambda w: w != "", contents):  # Filter out "empty" words.
                word = [char for char in word if char.isalnum() or char == "-"]  # Get all alphanumeric characters.
                # Join together all letters of the word again and make a list of words.
                all_words.append("".join(word))

        self.all_words = all_words

    def get_unique_words(self) -> set:
        """
        :return: a set of the unique words in the chosen file. Case insensitive.
        """

        words = []
        for word in self.all_words:
            words.append(word.lower())

        return set(words)

    def count_all_words(self) -> int:
        """
        :return: the count of all words in the chosen file.
        """

        return len(self.all_words)

    def count_word(self, word: str) -> int:
        """
        :param word: the word to count the occurrences of.
        :return: the count of the occurrences of word in the chosen file.
        """
        return self.all_words.count(word)

    def count_unique_words(self) -> int:
        """
        :return: the count of all unique words in the chosen file.
        """

        return len(self.get_unique_words())


class CommandLineInterface(cmd.Cmd):
    """
    Manage the command-line interface.
    """

    intro = "Welcome. I am uniQword, I can count all the words in your files and more.\n" \
            "To begin, select a file with the \"use\" command or type help or ? to read a list of commands."
    prompt = "uniQword, "
    file = None

    def no_file(self):
        """
        Alert the user that no file has been selected yet.
        """

        print("First I need to know which file to use!")
        self.onecmd("help use")

    @staticmethod
    def emptyline(**kwargs):
        """
        Scold the user for entering an empty command.
        """

        print("No idea what to do? Type help or ? to see a list of commands.")

    @staticmethod
    def default(line, **kwargs):
        """
        Scold the user for writing an unrecognised command.
        """

        print(f"I don't know of a command called \"{line}\". Please type help or ? to read a list of commands.")

    def do_use(self, file_path: str):
        """
        Select a file to operate on.
            Example: uniQword, use myfile.txt
        """

        try:
            self.file = WordsFile(file_path)

            print(f"I selected the file: {file_path}.")
        except FileNotFoundError:
            print("I couldn't find the file you asked for. Please try again.")
        except ValueError:
            print("I couldn't decode the given file. Please save it in UTF-8 or try another file.")

    def do_list(self, target: str):
        """
        Print a list of all requested elements (it can be long!). You can list: words, unique words.
            Examples:
                uniQword, list words
                uniQword, list uniques
        """

        if not self.file:
            self.no_file()
            return

        if not target:
            print("Please specify something to list!")
            self.onecmd("help list")
        elif target == "words":
            print(f"Here are all the words in the file:\n{self.file.all_words}")
        elif target in ["unique", "uniques", "unique words"]:
            print(f"Here are all the unique words in the file:\n{self.file.get_unique_words()}")

    def do_count(self, target: str):
        """
        Count how many of the specified element the file contains. You can count: *words, *unique words, specific words.
        To count specific words, write word, then the word you want to count.
            Examples:
                uniQword, count words
                uniQword, count unique words
                uniQword, count word banana
                uniQword, count w banana
        """

        if not self.file:
            self.no_file()
            return

        if target in ["*", "words"]:
            amount = self.file.count_all_words()
            print(f"The file contains: {amount:d} word{'' if amount == 1 else 's'} in total.")
        elif target in ["unique", "uniques", "unique words"]:
            amount = self.file.count_unique_words()
            print(f"The file contains: {amount:d} word{'' if amount == 1 else 's'} in total.")
        elif " " in target and target.split(" ")[0] in ["w", "word", "specific word"]:
            amount = self.file.count_word(target.split(" ")[1])
            print(f"The file contains: {amount:d} instance{'' if amount == 1 else 's'}"
                  f" of the word \"{target.split(' ')[1]}\".")
        else:
            print("Please specify something to count!")
            self.onecmd("help count")

    @staticmethod
    def do_bye(arg):
        """
        Exits the program.
            Example: uniQword, bye
        """

        del arg  # Apparently cmd requires me to accept an argument for every command, so I delete it. >:)
        print("See you, space cowboy!")
        time.sleep(2)
        exit()

    def do_test(self, arg):
        file = PyRTF.RTFFile(arg)
        print(file)


if __name__ == "__main__":
    CommandLineInterface().cmdloop()
