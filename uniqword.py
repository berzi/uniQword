"""uniQword is a program to read and count words from one or multiple files and perform some statistical operations."""

from typing import Optional  # Used for type hinting.
import cmd  # Used for the command-line interface.
import time  # Used by the command-line interface for sleep() when bidding farewell to the user.
import codecs  # Used to avoid codec problems when reading files.
import PyPDF2  # Used to read PDF files.
import docx  # Used to read docx files.
import zipfile  # This and next: used to read odt files.
from lxml import etree
import os  # Used for directory-wide operations.
# import shutil  # This and following: used to decode problematic passworded PDFs.
# import tempdir
# import subprocess
import collections  # Used for frequency counts.


# All currently accepted formats for files to examine.
SUPPORTED_FORMATS = (".txt", ".docx", ".odt", ".pdf")


class DecryptionError(Exception):
    """Catches the event in which an encrypted file is provided with a wrong password or none at all."""
    pass


class WordsFile:
    """Manage the file and collect and enumerate the words it contains."""
    file_words = []
    file_unique_words = set()
    file_path = ""

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

    def __repr__(self):
        """Represent the class as its own name plus the path of the contained file."""
        return f"{self.__class__.__name__}: {self.file_path}"

    def __bool__(self):
        """
        Test whether the WordsFile is not empty.
        :return: False if the file contains no words, True otherwise.
        """

        return len(self.file_words) > 0

    def __eq__(self, other):
        """Compare two instances on the base of the file path they point to."""
        return self.file_path == other.file_path

    def store_all_words(self):
        """
        Store an instance list with each word in the chosen file, eliminating every punctuation sign.
        :raise DecryptionError: if a wrong password was provided.
        :raise NotImplementedError: if the file is encrypted with an unsupported algorythm.
        :raise ValueError: if the provided file is of an unsupported format.
        """

        all_words = []
        contents = ""

        if self.file_path.endswith(".pdf"):
            with open(self.file_path, "rb") as pdf:
                reader = PyPDF2.PdfFileReader(pdf)  # Create a PDF handler.
                if reader.isEncrypted and self.password:
                    # Try to open the file with the given password.
                    # Will raise NotImplementedError if the algorythm is not supported by PyPDF2.
                    if reader.decrypt(self.password) == 0:
                        raise DecryptionError
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
        elif self.file_path.endswith(".txt"):
            with codecs.open(self.file_path) as file:
                contents = file.read()
        else:
            raise ValueError

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
            if len(word):
                # Remove hyphens at start or end.
                if word[0] == "-":
                    word.pop(0)
                if word[-1] == "-":
                    word.pop(-1)

                # Join together all letters of the word again and make a list of words.
                all_words.append("".join(word))

        self.file_words = all_words  # Store the list of words in an instance attribute for easy and cheap access.
        self.file_unique_words.update(all_words)  # Store all unique words in another attribute.

    def get_words(self) -> Optional[list]:
        """
        Get the list of the file's words.
        :return: the list of words or None.
        """

        if len(self.file_words):
            return self.file_words

        return None

    def get_unique_words(self) -> Optional[set]:
        """:return: a set of the unique words in the chosen file or None if no words are present."""
        if self.file_unique_words:
            return self.file_unique_words

        return None

    def count_all_words(self) -> int:
        """:return: the count of all words in the chosen file."""
        if self.words_count is None:
            self.words_count = len(self.file_words)

        return self.words_count

    def count_unique_words(self) -> int:
        """:return: the count of all unique words in the chosen file."""
        if self.uniques_count is None:
            self.uniques_count = len(self.get_unique_words())

        return self.uniques_count

    def count_word(self, word: str) -> int:
        """:return: the count of the occurrences of the specified word in the chosen file."""
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
    """

    files = {}  # Key: file name. Value: WordsFile instance.
    collective_words = []
    collective_unique_words = set()
    directories = {}  # Key: directory path. Value: list of file paths.

    # Attributes to act as a cache to optimise performance in case of repeated calls.
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

    def __repr__(self):
        """Represent the collection as its name plus a list of all files it contains."""
        if len(self.files):
            files = "\n".join([entry.__repr__() for entry in self.files.values()])
            return f"{self.__class__.__name__}:\n{files}"

        return f"{self.__class__.__name__}"

    def __str__(self):
        """Print the list of the files in the collection."""
        if len(self.files):
            return "\n".join([entry.__repr__() for entry in self.files.values()])

        return "The collection is empty."

    def __bool__(self):
        """:return: False if the collection is empty, True otherwise."""
        return len(self.files) > 0

    def __len__(self):
        """Return how many files the collection contains."""
        return len(self.files)

    def reset_values(self):
        """Reset all instance cache variables to force recounting all values."""
        self.collective_words_count = None
        self.collective_uniques_count = None
        self.collective_specific_count = {}
        self.collective_frequency_list = None

    def get_files(self) -> str:
        """Provide the file paths of each file in the collection."""

        for file_path in self.files.keys():
            yield file_path

    def add_files(self, *files: WordsFile):
        """
        Add the provided file(s) to the collection and their words to the collective words.
        :param files: one or more WordsFile to add to the collection.
        :raise TypeError: if the provided files are not valid WordsFile instances.
        """

        for file in files:
            if not isinstance(file, WordsFile):
                raise TypeError

            # Add the file to the collection using its file_path as index for optimal lookup.
            self.files.update({file.file_path: file})
            self.collective_words += file.get_words()
            self.collective_unique_words.update(file.get_words())

        self.reset_values()

    def remove_files(self, *file_paths: str) -> int:
        """
        Remove the provided files from the collection. File paths that are not found are ignored.
        :param file_paths: the file paths to remove from the collection.
        :raise ValueError: if no (valid) file is provided.
        :return: the number of files successfully deleted.
        """

        if not len(file_paths):
            raise ValueError("No file path to remove was provided.")

        removed = 0

        for file in file_paths:
            try:
                # Remove all words contained in the given file from the collection of words.
                for word in self.files[file].get_words():
                    self.collective_words.remove(word)
                    self.collective_unique_words.discard(word)

                del self.files[file]  # Delete the file itself from the collection.
                removed += 1
            except KeyError:
                continue  # Suppress the exception if no file with the given name is found.

        self.reset_values()
        return removed

    def add_directories(self, *directories: str) -> list:
        """
        Add the provided directory or directories to the collection by instantiating all files contained therein.
        Will not add files beginning in .
        :param directories: the path(s) of each directory to add.
        :raise ValueError: if no valid directory is provided.
        :return: the list of all files added successfully.
        """

        # TODO: optimise performance by packing/unpacking properly etc.
        if not len(directories):
            raise ValueError

        added = []
        for directory in directories:
            directory_files = []
            for file_name in os.listdir(directory):
                if file_name.endswith(SUPPORTED_FORMATS):
                    if file_name in self.files.keys():
                        continue  # Ignore files that are already in the collection.

                    try:
                        if directory != ".":
                            file_name = directory+"\\"+file_name

                        self.add_files(WordsFile(file_name, ""))

                        directory_files.append(file_name)
                        added.append(file_name)
                    except DecryptionError:
                        pass  # Suppress cases where passworded files are found, ignore them and move on.

            self.directories.update({directory: directory_files})

        return added

    def remove_directories(self, *directories: str) -> list:
        """
        Remove the provided directory or directories from the collection by removing each file contained therein.
        Files that had been added individually will not be removed even if they are present in the directory to remove.
        :param directories: the path(s) of each directory to remove.
        :raise ValueError: if no valid directory is provided.
        :return: the list of all files successfully removed.
        """

        # TODO: optimise performance by packing/unpacking properly etc.
        if not len(directories):
            raise ValueError

        removed = []
        for directory in directories:
            try:
                for file in self.directories[directory]:
                    self.remove_files(file)
                    removed.append(file)
            except KeyError:
                continue

        return removed

    def get_collective_words(self) -> Optional[list]:
        """:return: the list of all the files' words or None."""
        if len(self.collective_words):
            return self.collective_words

        return None

    def get_collective_unique_words(self) -> Optional[set]:
        """:return: a set of the unique words in the collection or None if no words are present."""
        if self.collective_unique_words:
            return self.collective_unique_words

        return None

    def count_collective_words(self) -> int:
        """:return: the count of all words in the collection."""
        if self.collective_words_count is None:
            self.collective_words_count = len(self.collective_words)

        return self.collective_words_count

    def count_collective_unique_words(self) -> int:
        """:return: the count of all unique words in the collection."""
        if self.collective_uniques_count is None:
            self.collective_uniques_count = len(self.collective_unique_words)

        return self.collective_uniques_count

    def count_collective_word(self, word: str) -> int:
        """:return: the count of the occurrences of the word in the collection."""
        return self.collective_specific_count.setdefault(word, self.collective_words.count(word))

    def get_frequency(self, top: int, reverse: bool = False) -> collections.Counter:
        """
        Get the frequency list of all words in the collection.
        :param top: the amount of words to return at most. Defaults to 10.
        :param reverse: whether the frequency list should show the least common items. Defaults to False.
        :return: a counter ["word"] = occurrences in descending order.
        """

        if not top:
            top = 10

        if self.collective_frequency_list is None:
            frequency_counter = collections.Counter()

            for word in self.collective_words:
                frequency_counter[word] += 1

            self.collective_frequency_list = frequency_counter.most_common()

        output = self.collective_frequency_list

        if output is None:
            return collections.Counter()

        if len(output) == 0:
            return collections.Counter()

        if reverse:
            output.reverse()

        output = output[:top]

        return output

    def print_stats(self, *, frequency_top: int = 10, frequency_reverse: bool = False) -> str:
        """
        Print all useful stats to a file.
        :return: the name of the file.
        """

        stats = "Stats for file"

        if len(self) == 1:
            # Use the only file's name as name for the stats file.
            file_name = "stats_"+self.files.__iter__().__next__()
            stats += ": {}".format(file_name)
        else:
            # Use uniQword as name for the stats file.
            file_name = "uniQword.txt"
            stats += "s:\n{}".format("\n".join(self.files.keys()))

        stats += "\n\n"
        # Add stats for word count.
        stats += "The {plural} {unique} unique words out of {total} total words.\n\n".format(
            plural="file contains" if len(self) == 1 else "files contain",
            unique=self.count_collective_unique_words(),
            total=self.count_collective_words()
        )

        # Add stats for frequency.
        words = self.get_frequency(top=frequency_top, reverse=frequency_reverse)
        output = []
        for word in words:
            # Base number of tabs minus the length of the word in tabs.
            tabs = "\t" * (6 - (len(word[0]) // 3))
            output.append(f"{word[0]}:{tabs}{word[1]}")

        stats += "{rev} frequent {amount} {plural}:\n{freq}".format(
            rev="Least" if frequency_reverse else "Most",
            amount=len(output),
            plural="word" if len(output) == 1 else "words",
            freq="\n".join(output)
        )

        with open(file_name+".txt", "w", encoding="UTF-8") as file:
            file.write(stats)

        return file_name


class CommandLineInterface(cmd.Cmd):
    """Manage the command-line interface."""
    intro = "Welcome. I am uniQword, I can count all the words in your files and more.\n" \
            "To begin, select a file with the \"add_file\" or \"add_dir\" command or type ? to read a list of commands."
    prompt = "uniQword, "
    file = FilesCollection()

    def check_file(self) -> bool:
        """:return: True if there is at least one  valid file selected, False otherwise."""
        if not self.file:
            print("I can't operate without a file!")
            self.onecmd("help add")
            return False

        if self.file.get_collective_words() is None:
            print(f"The selected file{'s are' if len(self.file) > 1 else ' is'} empty.")
            return False

        return True

    @staticmethod
    def emptyline(**kwargs):
        """Scold the user for entering an empty command."""
        print("No idea what to do? Type help or ? to see a list of commands.")

    @staticmethod
    def default(line, **kwargs):
        """Scold the user for writing an unrecognised command."""
        try:
            # Take only the first word if more are given.
            command = line.split(' ')[0]
        except IndexError:
            command = line
        print(f"I don't know of a command called \"{command}\".")
        CommandLineInterface.emptyline()

    def do_add_file(self, user_entry: str):
        """
        Select a file to operate on. You can select multiple items one at a time.
        Please provide a password if needed.
            Examples:
                uniQword, add_file myfile.pdf
                uniQword, add_file passwordedfile.pdf myp@ssw0rd
        """

        # TEST if adding in a nested dir works
        password = ""
        try:
            file, password = user_entry.split(" ")
        except ValueError:
            file = user_entry

        try:
            self.file.add_files(WordsFile(file, password))

            print(f"I selected the file: {user_entry}.")
        except FileNotFoundError:
            if len(file):
                print("I couldn't find the file you asked for. Please try again.")
            else:
                print("I need a file name in order to add it!")
                self.onecmd("help add")
        except ValueError:
            print("I couldn't decode the file. Please save it in UTF-8 before retrying.")
        except TypeError:
            print(f"I cannot use this file. Please convert it to one of the supported formats: "
                  f"{', '.join(SUPPORTED_FORMATS)}.")
        except DecryptionError:
            print("I need the correct password for this file!\n"
                  "Leave an empty space after the file name and type the password, example:\n"
                  "Qword, add myfile.txt myp@ssw0rd")
        except NotImplementedError:
            print("I couldn't decrypt the file. Please retry with a non-passworded copy.")

    def do_remove_file(self, target: str):
        """Remove a file from those currently in use. You can remove one file at a time or all at once.
            Examples:
                uniQword, remove_file myfile.txt
                uniQword, remove_file all files
                uniQword, remove_file *
        """

        if not self.file:
            print("There is no file to remove.")
            return

        if not target:
            print("Please select something to remove.")
            self.onecmd("help remove")
            return

        if target in ["all files", "*"]:
            removed = 0
            for file_path in self.file.get_files():
                removed += self.file.remove_files(file_path)

            print(f"I removed {'the only file' if removed == 1 else 'all '+str(removed)+' files'} from the list.")
        else:
            self.file.remove_files(target)
            print(f"I removed the file \"{target}\" if it was present in the list.")

    def do_add_dir(self, target: str):
        """
        Add all the files in a directory. Passworded or incompatible files will be ignored.
        Use "." to add the current directory.
            Examples:
                uniQword, add_dir myfolder
                uniQword, add_dir folder\myfolder
                uniQword, add_dir .
        """

        # TEST if adding a nested dir works
        try:
            print("I successfully added the following files:\n" +
                  "\n".join([file for file in self.file.add_directories(target)]))
        except ValueError:
            print("Please specify a folder to add!")
        except FileNotFoundError:
            print("I couldn't find the specified directory.")

    def do_remove_dir(self, target: str):
        """
        Remove all files from a previously added directory. Files that were added individually must be
        removed individually even if they are present in the directory to remove.
            Examples:
                uniQword, remove_dir myfolder
                uniQword, remove_dir folder\myfolder
                uniQword, remove_dir .
        """

        try:
            removed = self.file.remove_directories(target)
            if len(removed):
                print("I successfully removed the following files:\n" +
                      "\n".join([file for file in removed]))
            else:
                print("I couldn't remove the directory. Make sure you type the whole path correctly.\n"
                      "Do uniQword, files to check which file paths I'm currently using.")
        except ValueError:
            print("Please specify a folder to remove!\n"
                  "Do uniQword, files for a list of currently selected files and directories.")

    def do_files(self, arg):
        """
        List all files currently being processed.
        """

        del arg
        print("Here are all the files we're working on:\n"+"\n".join([entry for entry in self.file.get_files()]))

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
        words = self.file.get_collective_words()
        unique_words = self.file.get_collective_unique_words()

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
        else:
            print("Please specify something to list!")
            self.onecmd("help list")
            return

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
            amount = self.file.count_collective_words()
            print(f"The file contains: {amount:d} word{'' if amount == 1 else 's'} in total.")
        elif target in ["unique", "uniques", "unique words", "u"]:
            amount = self.file.count_collective_unique_words()
            print(f"The file contains: {amount:d} unique word{'' if amount == 1 else 's'} in total.")
        elif " " in target and target.split(" ")[0] in ["w", "word", "specific word"]:
            amount = self.file.count_collective_word(target.split(" ")[1])
            print(f"The file contains: {amount:d} instance{'' if amount == 1 else 's'}"
                  f" of the word \"{target.split(' ')[1]}\".")
        else:
            print("Please specify something to count!")
            self.onecmd("help count")

    def do_frequency(self, options: str= ""):
        """
        Print the frequency list of the current file. It can be printed in reverse and the maximum amount of results can
        be trimmed. By default, the first 10 results will be printed.
            Examples:
                uniQword, frequency
                uniQword, frequency reversed
                uniQword, frequency 15
                uniQword, frequency 15 reversed
        """

        if self.check_file() is False:
            return

        try:
            option_1, option_2 = options.split(" ")
        except ValueError:
            option_1 = options
            option_2 = None

        is_reversed = False
        top = None
        output = ""

        if option_1.strip().isnumeric():
            top = int(option_1.strip())
            if option_2 in ["r", "reverse", "reversed"]:
                is_reversed = True
        elif option_1 in ["r", "reverse", "reversed"]:
            is_reversed = True

        frequency = self.file.get_frequency(top=top, reverse=is_reversed)

        # Format the output to be easier on the eyes.
        for entry in frequency:
            # Calculate how many tabs to put in depending on the length of the word.
            tabs = "\t"*(6 - (len(entry[0])//4))
            output += f"{entry[0]}:{tabs}{entry[1]}\n"

        print(f"Here is the frequency list for the selected document(s){', reversed' if is_reversed else ''}:\n"
              f"{output}")

    def do_print(self, options):
        """
        Print all available stats on the currently selected document(s) to a file.
        The file will appear in the current folder and will be overwritten if already present.
        Options for the frequency list (number of items, reversed list) may be entered.
            Examples:
                uniQword, print
                uniQword, print reversed
                uniQword, print 15
                uniQword, print 15 reversed
                uniQword, print reversed 15
        """

        if self.check_file() is False:
            return

        try:
            option_1, option_2 = options.split(" ")
        except ValueError:
            option_1 = options
            option_2 = None

        is_reversed = False
        top = None

        if option_1.strip().isnumeric():
            top = int(option_1.strip())
            if option_2 in ["r", "reverse", "reversed"]:
                is_reversed = True
        elif option_1 in ["r", "reverse", "reversed"]:
            is_reversed = True
            if option_2.strip().isnumeric():
                top = int(option_2.strip())

        print(f"I printed data on {len(self.file)} file{'' if len(self.file) == 1 else 's'} on a file named "
              f"{self.file.print_stats(frequency_top=top, frequency_reverse=is_reversed)}.")

    @staticmethod
    def do_bye(arg):
        """
        Exit the program.
            Example: uniQword, bye
        """

        del arg  # Apparently cmd requires me to accept an argument for every command, so I delete it. >:)
        print("See you, space cowboy!")
        time.sleep(2)
        exit()

    def do_test(self, arg):
        """Test something."""
        pass


if __name__ == "__main__":
    CommandLineInterface().cmdloop()
