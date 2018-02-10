from typing import Optional
import cmd  # Used for the command-line interface.
import time  # Used by the command-line interface for sleep() when bidding farewell to the user.
import codecs  # Used to avoid codec problems when reading files.
import PyPDF2  # Used to read PDF files.
import docx  # Used to read docx files.
import zipfile  # Used to read odt files.
from lxml import etree  # This too.
# import os  # This and following: used to decode problematic passworded PDFs.
# import shutil
# import tempdir
# import subprocess
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

    # Attributes to optimise performance in case of repeated calls.
    words_count = None
    uniques_count = None
    specific_count = {}
    frequency_list = None

    password = ""

    def __init__(self, file_path: str, password: str):
        """
        Initialise the file instance by storing a list of all words.
        :param file_path: the file path and name.
        :param password: the password provided for the file, if given.
        """

        self.file_path = file_path
        if password:
            self.password = password

        self.store_all_words()

    def __bool__(self):
        """
        Test whether the WordsFile is not empty.
        :return: False if the file contains no words, True otherwise.
        """

        return len(self.file_words) > 0

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
                        # BUG: No idea why it seems to raise a FileNotFound exception.
                        # temporary_directory = tempdir.tempfile.mkdtemp(dir=os.path.dirname(self.file_path))
                        # temporary_pdf = os.path.join(temporary_directory, '_temp.pdf')

                        # subprocess.check_call(['qpdf', f"--password=", '--decrypt',
                        #                        self.file_path, temporary_pdf])

                        # shutil.move(temporary_pdf, self.file_path)

                        # Clean up the temporary dir.
                        # shutil.rmtree(temporary_directory)

                        raise NotImplementedError  # For now, just pass the exception for the UI to handle.
                elif reader.isEncrypted and not self.password:
                    raise DecryptionError

                for page in range(reader.numPages):
                    current_page = reader.getPage(page)  # Get one page at a time.
                    contents += current_page.extractText()  # Store the contents.
        elif self.file_path.endswith(".docx"):
            raw_document = docx.Document(self.file_path)
            for paragraph in raw_document.paragraphs:
                contents += paragraph.text
        elif self.file_path.endswith(".odt"):
            odt = zipfile.ZipFile(self.file_path)  # Open the file like a zip archive.
            # Get the content and take only the raw text.
            with odt.open('content.xml') as content:
                for child in etree.parse(content).iter():
                    if "text" in child.tag and child.text is not None:
                        # Add each tag's text to the context plus a line break to ensure words don't end up joined.
                        contents += child.text+"\n"
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
        for word in filter(lambda w: w not in ["", "\n"], contents):
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

    def get_words(self) -> Optional[list]:
        """
        Get the list of the file's words.
        :return: the list of words or None.
        """

        if len(self.file_words):
            return self.file_words

        return None

    def get_unique_words(self) -> Optional[set]:
        """
        :return: a set of the unique words in the chosen file or None if no words are present.
        """

        if self.get_words():
            return set(self.get_words())

        return None

    def count_all_words(self) -> int:
        """
        :return: the count of all words in the chosen file.
        """

        if self.words_count is None:
            self.words_count = len(self.file_words)

        return self.words_count

    def count_unique_words(self) -> int:
        """
        :return: the count of all unique words in the chosen file.
        """

        if self.uniques_count is None:
            self.uniques_count = len(self.get_unique_words())

        return self.uniques_count

    def count_word(self, word: str) -> int:
        """
        :param word: the word to count the occurrences of.
        :return: the count of the occurrences of word in the chosen file.
        """

        return self.specific_count.setdefault(word, self.file_words.count(word))

    def get_frequency(self) -> collections.Counter:
        """
        Get the frequency list of all words in the file.
        :return: a counter ["word"] = occurrences in descending order.
        """

        if self.frequency_list is None:
            frequency_counter = collections.Counter()

            for word in self.file_words:
                frequency_counter[word] += 1

            self.frequency_list = frequency_counter.most_common()

        return self.frequency_list


class FilesCollection:
    """
    Collect and manage all files to operate on.

    All functions are built to be compatible with the output of individual files (WordsFile).
    A FilesCollection can be treated as a WordsFile once initialised with at least one WordsFile.
    """

    files = {}
    collective_words = []

    # Attributes to optimise performance in case of repeated calls.
    collective_words_count = None
    collective_uniques_count = None
    collective_specific_count = {}
    collective_frequency_list = None

    def __init__(self, *files: Optional[WordsFile]):
        """
        Store all provided files.
        :param files: zero or more files to store.
        """

        if len(files) == 0:
            return

        self.add_files(*files)

    def reset_values(self):
        """
        Reset all instance variables to force recounting all values when the collection changes.
        """

        self.collective_words_count = None
        self.collective_uniques_count = None
        self.collective_specific_count = {}
        self.collective_frequency_list = None

    def get_files(self):
        """
        :return: a list of all the file paths currently contained in the collection.
        """

        pass  # TODO

    def add_files(self, *files: WordsFile):
        """
        Add the provided file(s) to the collection and their words to the collective words.
        :param files: one or more WordsFile to add to the collection.
        """

        for file in files:
            if not isinstance(file, WordsFile):
                raise TypeError

            # Add the file to the collection using its insertion order as index.
            self.files.update({len(self.files): file})
            self.collective_words.append(file.get_words())

        self.reset_values()

    def remove_files(self, *files):
        pass  # TODO: remember to reset_values()

    def get_collective_words(self) -> Optional[list]:
        """
        Get the list of all the files' words.
        :return: the list of words or None.
        """

        if len(self.collective_words):
            return self.collective_words

        return None

    def get_collective_unique_words(self) -> Optional[set]:
        """
        :return: a set of the unique words in the collection or None if no words are present.
        """

        if self.get_collective_words():
            return set(self.get_collective_words())

        return None

    def count_collective_words(self) -> int:
        """
        :return: the count of all words in the collection.
        """

        if self.collective_words_count is None:
            self.collective_words_count = len(self.collective_words)

        return self.collective_words_count

    def count_collective_unique_words(self) -> int:
        """
        :return: the count of all unique words in the collection.
        """

        if self.collective_uniques_count is None:
            self.collective_uniques_count = len(self.get_collective_unique_words())

        return self.collective_uniques_count

    def count_collective_word(self, word: str) -> int:
        """
        :param word: the word to count the occurrences of.
        :return: the count of the occurrences of word in the collection.
        """

        return self.collective_specific_count.setdefault(word, self.collective_words.count(word))

    def get_frequency(self) -> collections.Counter:
        """
        Get the frequency list of all words in the collection.
        :return: a counter ["word"] = occurrences in descending order.
        """

        if self.collective_frequency_list is None:
            frequency_counter = collections.Counter()

            for word in self.collective_words:
                frequency_counter[word] += 1

            self.collective_frequency_list = frequency_counter.most_common()

        return self.collective_frequency_list


class CommandLineInterface(cmd.Cmd):
    """
    Manage the command-line interface.
    """

    intro = "Welcome. I am uniQword, I can count all the words in your files and more.\n" \
            "To begin, select a file with the \"use\" command or type help or ? to read a list of commands."
    prompt = "uniQword, "
    file = None

    def check_file(self) -> bool:
        """
        Check if there's a file selected and if it contains any words.
        :return bool: True if file is valid, False otherwise.
        """

        if self.file is None:
            print("First I need to know which file to use!")
            self.onecmd("help use")
            return False

        if self.file.get_words() is None:
            print("The selected file is empty.")
            return False

        return True

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
            print("I couldn't decode the file. Please save it in UTF-8 before retrying.")
        except DecryptionError:
            print("I need the correct password for this file!\n"
                  "Leave an empty space after the file name and type the password, example:\n"
                  "Qword, use myfile.txt myp@ssw0rd")
        except NotImplementedError:
            print("I couldn't decrypt the file. Please use a non-passworded copy before retrying.")

    def do_list(self, target: str):
        """
        Print a list of all requested elements in the current file in no particular order (it can be long!).
        You can request: words, unique words.
            Examples:
                uniQword, list words
                uniQword, list unique words
                uniQword, list uniques
                uniQword, list w
                uniQword, list u
        """

        if self.check_file() is False:
            return

        output = ""
        words = self.file.get_words()
        unique_words = self.file.get_unique_words()

        if not target:
            print("Please specify something to list!")
            self.onecmd("help list")
            return

        if target in ["w", "words"]:
            line = ""
            for index, word in enumerate(words):
                line += word
                if len(line) >= 72:  # Limit the line length of each line to 72 chars.
                    line += "\n"
                    output += line
                    line = ""
                elif index == len(words) - 1:  # If we're done adding words.
                    output += line
                else:  # Add a comma.
                    line += ", "

            output = f"Here are all the words in the file:\n{output}"
        elif target in ["u", "unique", "uniques", "unique words"]:
            line = ""
            for index, word in enumerate(unique_words):
                line += word
                if len(line) >= 72:  # Limit the line length of each line to 72 chars.
                    line += "\n"
                    output += line
                    line = ""
                elif index == len(unique_words) - 1:  # If we're done adding words.
                    output += line
                else:  # Add a comma.
                    line += ", "

            output = f"Here are all the unique words in the file:\n{output}"

        print(output)

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

        if self.check_file() is False:
            return

        if target in ["*", "words", "w"]:
            amount = self.file.count_all_words()
            print(f"The file contains: {amount:d} word{'' if amount == 1 else 's'} in total.")
        elif target in ["unique", "uniques", "unique words", "u"]:
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

        if self.check_file() is False:
            return

        is_reversed = False
        frequency = self.file.get_frequency()
        output = ""

        # Reverse the list if necessary.
        if order in ["r", "reverse", "reversed"]:
            is_reversed = True
            frequency = sorted(list(frequency), key=operator.itemgetter(1))

        # Format the output to be easier on the eyes.
        for entry in frequency:
            # Calculate how many tabs to put in depending on the length of the word.
            tabs = "\t" * (5 - (len(entry[0])+1) // 4)
            output += f"{entry[0]}:{tabs}{entry[1]}\n"

        print(f"Here is the frequency list of the file{', reversed' if is_reversed else ''}:\n"
              f"{output}")

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
