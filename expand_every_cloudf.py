# -*- coding: utf-8 -*-
"""Traverse a directory recursively until no .cloudf folder placeholder left.

[odrive](https://www.odrive.com) creates placeholder files for directories and files
stored in the cloud so they don't use any disk space unless you try to open them,
when they are synced to the disk.
They have the .cloud extension for files and .cloudf for folders.

This whole expansion costs 0 byte disk space but gives you access to all your
cloud files either directly or as a .cloud placeholder.

If the emojis aren't printed properly on Windows you can try the
[Windows Terminal](https://github.com/microsoft/terminal).
"""
import argparse
import multiprocessing as mp
import os
import re
import sys
import time


def parse_args():
    """Return the parsed cli arguments."""
    parser = argparse.ArgumentParser(description="Expand .cloudf files into actual directories")
    parser.add_argument("-d", "--directory", help="The root directory where our journey begins at.")
    parser.add_argument("-e", "--exclude-pattern",
                        help="A regular expression which is searched in the absolute path of a cloudf file.\n"
                             "In case of a match, the cloudf file won't get expanded.")
    parser.add_argument("-t", "--threads", type=int, default=None,
                        help="How many threads the script should use to expand folders.\n"
                             "Default: number of CPU cores")
    args = parser.parse_args()

    if args.directory is None:
        args.directory = input("Enter the root directory: ")
        args.directory = args.directory.replace('"', '')

    return args


def expand(cloud_folder_full_path):
    """Double-click on the folder on the `cloud_folder_full_path` path."""
    # TODO does this work on Linux?
    os.startfile(cloud_folder_full_path)


def try_expanding(cloud_folder_full_path, sleep_between_tries=0.1, timeout=10 * 60):
    start_time = time.time()
    while os.path.exists(cloud_folder_full_path):
        try:
            expand(cloud_folder_full_path)
            time.sleep(sleep_between_tries)
        except FileNotFoundError:
            pass

        elapsed_time = time.time() - start_time
        if elapsed_time > timeout:
            return False
    return True


def traverse_and_expand(directory, exclude_pattern, num_threads):
    """Open every cloudf folder until none left under the user specified `directory`."""
    if num_threads is None:
        num_threads = os.cpu_count()
    print(f"using {num_threads} threads for folder expansion")

    with mp.Pool(num_threads) as pool:
        there_might_be_more = True
        cycle_counter = 0
        while there_might_be_more:
            async_results = []

            cycle_counter += 1
            last_progress_msg_len = 0
            print(f"traversing #{cycle_counter}...")
            for (root, _, filenames) in os.walk(directory):
                progress_msg_len = len(root)
                print(' ' * last_progress_msg_len, end='\r')
                print(root, end='\r')
                last_progress_msg_len = progress_msg_len

                cloud_folder_names = [filename for filename in filenames
                                      if re.search(r"\.cloudf$", filename) is not None]
                for cloud_folder_name in cloud_folder_names:
                    cloud_folder_full_path = os.path.join(root, cloud_folder_name)

                    if (exclude_pattern is not None) and re.search(exclude_pattern, cloud_folder_full_path):
                        print(f"🛑✋ exclude {cloud_folder_full_path} ✋🛑")
                        pass
                    else:
                        print(f"🟢👍 expand {cloud_folder_full_path} ... 👍🟢")
                        async_result = pool.apply_async(try_expanding, args=(cloud_folder_full_path,))
                        async_results.append((cloud_folder_full_path, async_result))

            print(' ' * last_progress_msg_len, end='\r')
            print(f"wait for the expansion threads #{cycle_counter}...")
            for (path, async_result) in async_results:
                print(f"fetching the result for {path}... ", end='')
                sys.stdout.flush()
                result = "✔" if async_result.get() else "❌"
                print(result)

            there_might_be_more = len(async_results) != 0


def traverse_and_open_all():
    """Parse args and then start the traversal on a single or multiple threads."""
    args = parse_args()

    traverse_and_expand(args.directory, args.exclude_pattern, args.threads)

    print("🎉🎉🎉 All done! 🎉🎉🎉")


if __name__ == "__main__":
    traverse_and_open_all()
