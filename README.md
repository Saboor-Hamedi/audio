# Convert to Audio

A command-line tool to convert a `.docx` file to an audio file.

## Installation

Follow these steps to install and use the `convert-to-audio` package.

### 1. **Download the Package**

You can either download the package from a distribution file or install it directly from PyPI.

#### Option A: From Distribution File

1. **Download the Distribution Files**:
   - Obtain the `.whl` or `.tar.gz` file from the distribution.

2. **Install the Package**:
   - Open your terminal or command prompt.
   - Navigate to the directory containing the downloaded file.
   - Run the following command to install the package:

     ```sh
     pip install path_to_the_file.whl
     ```

     or

     ```sh
     pip install path_to_the_file.tar.gz
     ```

#### Option B: From PyPI

1. **Install via PyPI**:
   - Open your terminal or command prompt.
   - Run the following command to install the package directly from PyPI:

     ```sh
     pip install convert-to-audio
     ```

### 2. **Usage**

After installation, you can use the `convert-to-audio` command to convert a `.docx` file to an audio file. 

Here’s the syntax for using the command:

```sh
convert-to-audio <input_file> <output_file> <voice_id> <rate> <volume>
