"""uniQword is a program to read and count words from one or multiple files and perform some statistical operations."""

import cmd  # Used for the command-line interface.
import codecs  # Used to avoid codec problems when reading files.
import collections  # Used for frequency counts.
import os  # Used for directory-wide operations.
import time  # Used by the command-line interface for sleep() when bidding farewell to the user.
import zipfile  # Used to read odt files.
import re  # Used for text parsing.
from typing import Optional  # Used for type hinting.

import PyPDF2  # Used to read PDF files.
import docx  # Used to read docx files.
from lxml import etree  # Used to read odt files.

# All currently accepted formats for files to examine.
SUPPORTED_FORMATS = (".txt", ".docx", ".odt", ".pdf")

# The default number of elements for frequency lists if unspecified by user input.
FREQUENCY_TOP = 20

# Symbols to accept within words.
ACCEPT = ("-", "_")

# Subset of ACCEPT to remove from start/end of words.
REMOVE = ("-",)

# Symbols (regex) to count as word separators.
SEPARATORS = r"\s'"


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

        all_words = self.purify_words(contents)

        if len(all_words):  # Avoid adding empty contents.
            self.file_words = all_words  # Store the list of words in an instance attribute for easy and cheap access.
            self.file_unique_words.update(all_words)  # Store all unique words in another attribute.

    @staticmethod
    def purify_words(contents: str) -> list:
        """
        Clean up words by removing empty words, whitespace and symbols.
        :param contents: the string to purify.
        :return: a list of purified words.
        """

        all_words = []

        # Separate words.
        contents = re.split(r"["+SEPARATORS+r"]", contents)

        # Filter out "empty" words and filter characters inside words to make sure we only get real(istic) words.
        for word in filter(lambda w: w not in ["", "\n"], contents):
            # Get all alphanumeric characters, plus hyphens and underscores.
            word = [char for char in word if char.isalnum() or char in ACCEPT]
            while len(word):  # Ensure we're not working on an empty word.
                # Remove hyphens at start or end.
                if word[0] in REMOVE:
                    word.pop(0)
                    continue
                if word[-1] in REMOVE:
                    word.pop(-1)
                    continue

                # Join together all letters of the word again and make a list of words.
                all_words.append("".join(word))
                break

        return all_words

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
            finally:
                # Clean up the cache if at least one file was correctly deleted.
                if removed > 0:
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

    def get_frequency(self, top: int=FREQUENCY_TOP, reverse: bool = False) -> collections.Counter:
        f"""
        Get the frequency list of all words in the collection.
        :param top: the amount of words to return at most. Defaults to {str(FREQUENCY_TOP)}. 0 outputs the whole list.
        :param reverse: whether the frequency list should show the least common items. Defaults to False.
        :return: a counter ["word"] = occurrences in descending order.
        """

        if top is None:
            top = FREQUENCY_TOP

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

        if top == 0:
            top = len(output)

        output = output[:top]

        return output

    def print_stats(self, *, frequency_top: int=0, frequency_reverse: bool = False) -> str:
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

        # Stuff for string padding.
        longest_word = max([len(word[0]) for word in words])
        if longest_word + 4 <= 60:
            longest_word += 4
        else:
            longest_word = 60

        # Format the output to be easier on the eyes.
        for entry in words:
            # Calculate how many tabs to put in depending on the length of the word.
            output.append(f"{entry[0]:{longest_word}}{entry[1]}")

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
            "To begin, select a file with the \"add\" command or type ? to read a list of commands."
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

    def do_add(self, user_entry: str):
        """
            Select a file or directory to operate on. You can select multiple items one at a time.
            Please provide a password if needed. Passworded files will be ignored when adding an entire directory.
            To add all compatible files in the current directory, type .
                Examples:
                    uniQword, add .
                    uniqword, add mydir\folder
                    uniQword, add myfile.pdf
                    uniQword, add passwordedfile.pdf myp@ssw0rd
        """

        # Identify if the user provided any input at all.
        if not user_entry:
            print("Plase specify something to add! Type \"uniQword, help add\" to receive help.")
            return

        user_entry.strip()  # Remove trailing spaces.

        # Identify if the user asked for a directory or a file.
        if user_entry == "." or "." not in user_entry:
            try:
                added = [file for file in self.file.add_directories(user_entry)]
                if len(added):
                    print(f"I successfully added the following file{'' if len(added) == 1 else 's'}:\n" +
                          "\n".join(added))
                else:
                    print("I couldn't find any compatible file. Please remember to add passworded files individually.")
            except FileNotFoundError or NotADirectoryError:
                print("I couldn't find the specified directory.")
        else:
            try:
                file, password = user_entry.split(" ")
            except ValueError:
                file = user_entry
                password = ""

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
                      "uniQword, add myfile.txt myp@ssw0rd")
            except NotImplementedError:
                print("I couldn't decrypt the file. Please retry with a non-passworded copy.")

    def do_remove(self, user_entry: str):
        """
        Remove a file or directory from use.
        Use . to remove the current directory only, and * to remove all files and directory.
        Type "uniQword, files" to read a list of the files currently in use.
            Examples:
                uniQword, remove .
                uniQword, remove myfolder
                uniQword, remove mydir\myfolder
                uniQword, remove myfile.txt
                uniQword, remove *
        """

        if not self.file:
            print("There are no files to remove.")
            return

        if not user_entry:
            print("Please select something to remove.")
            self.onecmd("help remove")
            return

        user_entry.strip()

        # Check if the user wants to clear the list.
        if user_entry in ["*"]:
            removed = 0
            for file_path in self.file.get_files():
                removed += self.file.remove_files(file_path)

            print(f"I removed {'the only file' if removed == 1 else 'all '+str(removed)+' files'} from the list.")
            return

        # Try to remove a file.
        try:
            self.file.remove_files(user_entry)
            print(f"I removed the file \"{user_entry}\" from the list.")
        except ValueError:
            # If it doesn't work, it may be a directory.
            try:
                removed = self.file.remove_directories(user_entry)
                if len(removed):
                    print(f"I successfully removed the following file{'' if len(removed) == 1 else 's'}:\n" +
                          "\n".join([file for file in removed]))
                else:
                    print("I couldn't remove any file. Make sure you type the whole path correctly.\n"
                          "Do \"uniQword, files\" to check which file paths I'm currently using.")
            except ValueError:
                print("Please specify a valid file or folder to remove!\n"
                      "Do \"uniQword, files\" for a list of currently selected files and directories.")

    def do_files(self, arg):
        """List all files currently being processed."""

        del arg
        print("Here are all the files we're working on:\n"+"\n".join([entry for entry in self.file.get_files()]))

    def do_count(self, user_entry):
        """
        Count the number of words and unique words in the currently selected files.
        Or count how many times a specific word occurs in the currently selected files.
            Example:
                uniQword, count
                uniQword, count banana
        """

        if self.check_file() is False:
            return

        user_entry.strip()

        # Check if the user wants to count a specific word.
        if user_entry:
            occurrences = self.file.count_collective_word(user_entry)

            print(f"The file{'' if len(self.file) == 1 else 's'} contain{'' if len(self.file) == 1 else 's'} "
                  f"{occurrences} occurrences of the word \"{user_entry}\".")
            return

        total_words = self.file.count_collective_words()
        total_uniques = self.file.count_collective_unique_words()

        print(f"The file{'' if len(self.file) == 1 else 's'} contain{'' if len(self.file) == 1 else 's'} "
              f"{total_words} total words, {total_uniques} of which unique "
              f"({round((total_uniques / total_words) * 100, 2)}%).")

    def do_frequency(self, options: str):
        f"""
        Print the frequency list of the current file. It can be printed in reverse and the maximum amount of results can
        be trimmed. By default, the first {str(FREQUENCY_TOP)} results will be printed. Input * to print all results.
            Examples:
                uniQword, frequency
                uniQword, frequency *
                uniQword, frequency * reversed
                uniQword, frequency reversed
                uniQword, frequency 50
                uniQword, frequency 50 reversed
        """

        if self.check_file() is False:
            return

        try:
            option_1, option_2 = options.split(" ")
            option_2.strip()
        except ValueError:
            option_1 = options
            option_2 = None

        is_reversed = False
        top = None
        output = ""
        option_1.strip()

        if option_1.isnumeric():
            top = int(option_1.strip())
            if option_2 in ["r", "reverse", "reversed"]:
                is_reversed = True
        elif "*" in option_1:
            top = 0
            if option_2 in ["r", "reverse", "reversed"]:
                is_reversed = True
        elif option_1 in ["r", "reverse", "reversed"]:
            is_reversed = True
            if "*" in option_2:
                top = 0

        frequency = self.file.get_frequency(top=top, reverse=is_reversed)

        # Stuff for string padding.
        longest_word = max([len(word[0]) for word in frequency])
        if longest_word+4 <= 60:
            longest_word += 4
        else:
            longest_word = 60

        # Format the output to be easier on the eyes.
        for entry in frequency:
            # Calculate how many tabs to put in depending on the length of the word.
            output += f"{entry[0]:{longest_word}}{entry[1]}\n"

        print(f"Here are the {'least' if is_reversed else 'most'} common {str(top) if top else str(FREQUENCY_TOP)} "
              f"elements for the selected document{'' if len(self.file) == 1 else 's'}:\n{output}")

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
            option_2.strip()
        except ValueError:
            option_1 = options
            option_2 = None

        is_reversed = False
        top = 0
        option_1.strip()

        if option_1.isnumeric():
            top = int(option_1.strip())
            if option_2 in ["r", "reverse", "reversed"]:
                is_reversed = True
        elif option_1 in ["r", "reverse", "reversed"]:
            is_reversed = True

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
