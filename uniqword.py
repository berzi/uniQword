import cmd  # Used for the command-line interface.
import time  # Used by the command-line interface for sleep() when bidding farewell to the user.
import codecs  # Used to avoid codec problems when reading files.
import PyPDF2  # Used to read PDF files.
import docx  # Used to read docx files.
import os  # This and following: used to decode problematic passworded PDFs.
import shutil
import tempdir
import subprocess
import collections  # Used for frequency counts.
import operator  # Used for itemgetter() to optimise reverse operations.


class DecryptionError(Exception):
    """
    Catches the event in which an encrypted file is provided with a wrong password or none at all.
    """

    pass


class WordsFile:
    """
    Manage the file and collect and enumerate the words it contains.
    """

    file_words = []
    password = ""

    def __init__(self, file_path: str, password: str):
        """
        Initialise the file instance by storing a list of all words.
        :param file_path: the file path and name.
        """

        self.file_path = file_path
        if password:
            self.password = password

        self.store_all_words()

    def store_all_words(self):
        """
        Store an instance list with each word in the chosen file, eliminating every punctuation sign.
        """

        all_words = []
        contents = ""

        if self.file_path.endswith(".pdf"):
            with open(self.file_path, "rb") as pdf:
                reader = PyPDF2.PdfFileReader(pdf)  # Create a PDF handler.
                if reader.isEncrypted and self.password:
                    try:
                        # Try to open the file with the given password.
                        if reader.decrypt(self.password) == 0:
                            # If the password is wrong, raise an exception for the UI to handle.
                            raise DecryptionError
                    except NotImplementedError:
                        # If PyPDF can't handle the encryption algorithm, do it with a subprocess call.
                        # Thanks to GitHub user ssokolow for this bit.
                        # Make a temporary directory and file to work with safely.
                        # TODO: No idea why it seems to raise a FileNotFound exception.
                        temporary_directory = tempdir.tempfile.mkdtemp(dir=os.path.dirname(self.file_path))
                        temporary_pdf = os.path.join(temporary_directory, '_temp.pdf')

                        subprocess.check_call(['qpdf', f"--password=", '--decrypt',
                                               self.file_path, temporary_pdf])

                        shutil.move(temporary_pdf, self.file_path)

                        # Clean up the temporary dir.
                        shutil.rmtree(temporary_directory)
                elif reader.isEncrypted and not self.password:
                    raise DecryptionError

                for page in range(reader.numPages):
                    current_page = reader.getPage(page)  # Get one page at a time.
                    contents += current_page.extractText()  # Store the contents
        elif self.file_path.endswith(".docx"):
            raw_document = docx.Document(self.file_path)
            for paragraph in raw_document.paragraphs:
                contents += paragraph.text
        else:
            with codecs.open(self.file_path) as file:
                contents = file.read()

        # Get all to lowercase. This is useful to count unique words, and we don't need case sensitiveness anyway.
        contents = contents.lower()

        # Get individual lines.
        contents = contents.splitlines()

        if self.file_path.endswith(".pdf"):
            # Eliminate the fake line breaks PDFs have.
            contents = "".join(contents)
            # Real line breaks in PDFs automatically get some whitespace, so we don't need to join words using it.
        else:
            # Join all words and lines with whitespace, which we'll use to separate individual words.
            contents = " ".join(contents)

        # Separate words.
        contents = contents.split(" ")

        # Filter out "empty" words and filter characters inside words to make sure we only get real(istic) words.
        for word in filter(lambda w: w != "", contents):
            # Get all alphanumeric characters, plus hyphens and underscores.
            word = [char for char in word if char.isalnum() or char in ["-", "_"]]

            # Remove hyphens at start or end.
            if word[0] == "-":
                word.pop(0)
            if word[-1] == "-":
                word.pop(-1)

            # Join together all letters of the word again and make a list of words.
            all_words.append("".join(word))

        self.file_words = all_words  # Store the list of words in an instance attribute for easy and cheap access.

    def get_unique_words(self) -> set:
        """
        :return: a set of the unique words in the chosen file.
        """

        return set(self.file_words)

    def count_all_words(self) -> int:
        """
        :return: the count of all words in the chosen file.
        """

        return len(self.file_words)

    def count_word(self, word: str) -> int:
        """
        :param word: the word to count the occurrences of.
        :return: the count of the occurrences of word in the chosen file.
        """
        return self.file_words.count(word)

    def count_unique_words(self) -> int:
        """
        :return: the count of all unique words in the chosen file.
        """

        return len(self.get_unique_words())

    def get_frequency(self) -> collections.Counter:
        """
        Get  the frequency list of all words in the file.
        :return: a counter ["word"] = occurrences in descending order.
        """

        frequency_counter = collections.Counter()

        for word in self.file_words:
            frequency_counter[word] += 1

        return frequency_counter.most_common()


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

        print(f"I don't know of a command called \"{line}\".")
        CommandLineInterface.emptyline()

    def do_use(self, user_entry: str):
        """
        Select a file to operate on.
            Example: uniQword, use myfile.txt
        """

        password = ""
        try:
            file, password = user_entry.split(" ")[0], user_entry.split(" ")[1]
        except IndexError:
            file = user_entry

        try:
            self.file = WordsFile(file, password)

            print(f"I selected the file: {user_entry}.")
        except FileNotFoundError:
            print("I couldn't find the file you asked for. Please try again.")
        except ValueError:
            print("I couldn't decode the given file. Please save it in UTF-8 or try another file.")
        except DecryptionError:
            print("I need the correct password for this file!\n"
                  "Leave an empty space after the file name and type the password, example:\n"
                  "Qword, use myfile.txt myp@ssw0rd")

    def do_list(self, target: str):
        """
        Print a list of all requested elements in the current file (it can be long!).
        You can request: words, unique words.
            Examples:
                uniQword, list words
                uniQword, list unique words
                uniQword, list uniques
        """

        if not self.file:
            self.no_file()
            return

        if not target:
            print("Please specify something to list!")
            self.onecmd("help list")
        elif target in ["w", "words"]:
            print(f"Here are all the words in the file:\n"
                  f"{self.file.file_words}")
        elif target in ["u", "unique", "uniques", "unique words"]:
            print(f"Here are all the unique words in the file:\n"
                  f"{self.file.get_unique_words()}")

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

    def do_frequency(self, order: str=""):
        """
        Print the frequency list of the current file. Can be printed in reverse.
            Examples:
                uniQword, frequency
                uniQword, frequency reversed
        """

        if not self.file:
            self.no_file()
            return
        is_reversed = False
        frequency = self.file.get_frequency()

        if order in ["r", "reverse", "reversed"]:
            is_reversed = True
            frequency = sorted(list(frequency), key=operator.itemgetter(1))

        print(f"Here is the frequency list of the file{', reversed' if is_reversed else ''}:\n"
              f"{frequency}")

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
        pass


if __name__ == "__main__":
    CommandLineInterface().cmdloop()
