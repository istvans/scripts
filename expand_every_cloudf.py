# -*- coding: utf-8 -*-
"""Traverse a directory recursively until no cloudf placeholders left.

[odrive](https://www.odrive.com) creates placeholder files for directories and files
stored in the cloud so they don't use any disk space unless you try to open them,
when they are synced to the disk.
They have the .cloud extension of files and .cloudf for folders.

This whole expansion costs 0 byte disk space.

If the emojis aren't printed properly on Windows you can try the
[Windows Terminal](https://github.com/microsoft/terminal).
"""
import argparse
import itertools
import math
import os
import re
import sys
import time


class Spinner():
    r"""Spinning single character progress presenter with built-in timeout and retry mechanisms.

      \|/
    -- * --
      /|\
    """

    class TimeoutException(Exception):
        """Raised when we are still spinning after the specified timeout."""
        pass

    def __init__(self, timeout, spin_time=0.1, retry_fn=None, retry_time=60):
        assert isinstance(timeout, int)
        assert timeout > 0
        self._spinner = itertools.cycle(["-", "\\", "|", "/"])
        self.single_spin_time = spin_time
        self.timeout = timeout
        self.retry_fn = retry_fn
        self.retry_time = retry_time
        self.retry_spins = math.ceil(self.retry_time / self.single_spin_time)
        self.timeout_spins = math.ceil(self.timeout / self.single_spin_time)
        self.num_spins = 0

    def _raise_or_retry(self):
        """Check how long we've been spinning so far and do what's necessary (might be nothing)."""
        if self.num_spins != 0:
            if self.num_spins >= self.timeout_spins:
                approx_elapsed_sec = self.num_spins * self.single_spin_time
                raise Spinner.TimeoutException(f"Still spinning after {self.num_spins}: "
                                               f"~{approx_elapsed_sec}s >= {self.timeout}s")
            elif self.retry_fn is not None and (self.num_spins % self.retry_spins) == 0:
                self.retry_fn()
            else:
                pass  # normal spinning

    def spin(self):
        """Print one cycle of the spinning then wait `for_sec`"""
        self._raise_or_retry()
        sys.stdout.write(next(self._spinner))
        sys.stdout.flush()
        sys.stdout.write("\b")
        time.sleep(self.single_spin_time)
        self.num_spins += 1

    def __enter__(self):
        """Let's start spinning!"""
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Remove the last printed spinner character."""
        pass


def item_level_string(tree_level):
    """Return a string representation of the specified `tree_level` (aka depth)."""
    return tree_level * "---"


def parse_args():
    """Return the parsed cli arguments."""
    parser = argparse.ArgumentParser(description="Expand .cloudf files into actual directories")
    parser.add_argument("-d", "--directory", help="The root directory where our journey begins at.")
    parser.add_argument("-e", "--exclude-pattern",
                        help="A regular expression which is searched in the absolute path of a cloudf file.\n"
                             "In case of a match, the cloudf file won't get expanded.")
    args = parser.parse_args()

    if args.directory is None:
        args.directory = input("Enter the root directory: ")
        args.directory = args.directory.replace('"', '')

    return args


def clear_screen():
    """Clear the command line screen."""
    os.system('cls' if os.name == 'nt' else 'clear')


def double_click(file_path):
    """Double-click the file on the `file_path`."""
    # TODO does this work on Linux?
    os.startfile(file_path)


def traverse_and_open_all():
    """Open every cloudf folder and traverse recursively until none left under the user specified directory."""
    args = parse_args()

    cloudf_was_found = True
    round_ = 0
    while cloudf_was_found:
        clear_screen()
        round_ += 1
        print("Round", round_)
        cloudf_was_found = False
        for (root, _, filenames) in os.walk(args.directory):
            path = root.split(os.sep)
            level = len(path) - 1

            depth_str = item_level_string(level)
            print(depth_str, os.path.basename(root))

            cloud_folder_names = [filename for filename in filenames if ".cloudf" in filename]
            for cloud_folder_name in cloud_folder_names:
                depth_str = item_level_string(level + 1)

                cloud_folder_full_path = os.path.join(root, cloud_folder_name)
                if (args.exclude_pattern is not None) and re.search(args.exclude_pattern, cloud_folder_full_path):
                    print(depth_str, f"🛑✋ NOT expanding {cloud_folder_name}! ✋🛑")
                    continue
                else:
                    cloudf_was_found = True
                    print(depth_str, f"expanding {cloud_folder_name}... ", end='')

                double_click(cloud_folder_full_path)

                spinner = Spinner(timeout=20 * 60,  # 20 minutes
                                  retry_fn=lambda: double_click(cloud_folder_full_path))
                try:
                    with spinner:
                        while os.path.exists(cloud_folder_full_path):
                            spinner.spin()
                except Spinner.TimeoutException:
                    print("❌")
                else:
                    print("✔")
    print("🎉🎉🎉 All done! 🎉🎉🎉")


if __name__ == "__main__":
    traverse_and_open_all()
