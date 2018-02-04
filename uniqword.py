import cmd
import time
import codecs


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
                word = [char for char in word if char.isalnum()]  # Get all alphanumeric characters.
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

    def emptyline(self):
        """
        Scold the user for entering an empty command.
        """

        print("No idea what to do? Type help or ? to see a list of commands.")

    def default(self, line):
        """
        Scold the user for writing an unrecognised command.
        """

        print("I don't of a command called \"%s\". Please type help or ? to read a list of commands." % line)

    def do_use(self, file_path: str):
        """
        Select a file to operate on.
            Example: uniQword, use myfile.txt
        """

        try:
            self.file = WordsFile(file_path)

            print("I selected the file: %s." % file_path)
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

        if target == "":
            print("Please specify something to list!")
            self.onecmd("help list")
        elif target == "words":
            print("Here are all the words in the file:\n%s" % self.file.all_words)
        elif target in ["unique", "uniques", "unique words"]:
            print("Here are all the unique words in the file:\n%s" % self.file.get_unique_words())

    def do_count(self, target: str):
        """
        Count how many of the specified element the file contains. You can count: *words, *unique words, specific words.
        The * asterisks are important!
            Examples:
                uniQword, count *words
                uniQword, count *unique words
                uniQword, count banana
        """

        if not self.file:
            self.no_file()
            return

        if target == "":
            print("Please specify something to count!")
            self.onecmd("help count")
        elif target in ["*", "*words"]:
            amount = self.file.count_all_words()
            plural = "s"
            if amount == 1:
                plural = ""
            print("The file contains: %i word%s in total." % (amount, plural))
        elif target in ["*unique", "*uniques", "*unique words"]:
            amount = self.file.count_unique_words()
            plural = "s"
            if amount == 1:
                plural = ""
            print("The file contains: %i unique word%s in total." % (amount, plural))
        else:
            amount = self.file.count_word(target)
            plural = "s"
            if amount == 1:
                plural = ""
            print("The file contains: %i instance%s of the word \"%s\"." % (amount, plural, target))

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


if __name__ == "__main__":
    CommandLineInterface().cmdloop()
