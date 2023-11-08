"""
    Module containing functions to perform common
    file/folder handling operations.
"""

import os
from tkinter import Tk, filedialog
from os import listdir, makedirs
from os.path import join, split, exists, isdir
import cv2
import constants


def get_folder():
    """
    Function to use filedialog to allow user select a folder on their computer

    Returns
    -------
    str
        absolute path to selected folder
    """

    root = Tk()  # Pointing root to Tk() to use it as Tk() in program.
    root.withdraw()  # Hides small tkinter window.
    root.attributes("-topmost", True)  # Opened window will be active and above all
    # windows despite of selection.
    folder = filedialog.askdirectory()  # Open file dialog that returns selected
    # absolute paths as stringsr
    return folder


def get_file():
    """
    Function to use filedialog to allow user select a file on their computer

    Returns
    -------
    file_path : str
        absolute file path to selected file
    """

    root = Tk()  # Pointing root to Tk() to use it as Tk() in program.
    root.withdraw()  # Hides small tkinter window.
    root.attributes("-topmost", True)  # Opened window will be active and above all
    # windows despite of selection.
    filename = filedialog.askopenfilename()  # Open file dialog that returns selected
    # absolute paths as stringsr
    return filename


def get_files():
    """
    Function to use filedialog to allow user select multiple files on their computer

    Returns
    -------
    file_paths : str iterable
        absolute file paths to selected files
    """

    root = Tk()  # Pointing root to Tk() to use it as Tk() in program.
    root.withdraw()  # Hides small tkinter window.
    root.attributes("-topmost", True)  # Opened window will be active
    # and above all windows despite of selection.
    filenames = filedialog.askopenfilenames()  # Open file dialog that
    # returns selected absolute paths as stringsr
    return filenames


def get_davis_set_files_from_folder(folder_path):
    """
    Function to get files from directory that contain the davis set extension

    Parameters
    ----------
    folder_path : str
        directory to extract files from

    Returns
    -------
    file_paths : str iterable
        absolute file path to files ending in .set
    """

    return get_files_from_folder(
        folder_path, only_files_with_extension=constants.DAVIS_SET_FILE_EXTENSION
    )


def get_excel_files_from_folder(folder_path):
    """
    Function to get file paths to files in directory that contain the davis set extension

    Parameters
    ----------
    folder_path : str
        directory to extract files from

    Returns
    -------
    file_paths : str iterable
        absolute file path to files ending in excel extension
    """

    return get_files_from_folder(
        folder_path, only_files_with_extension=constants.EXCEL_FILE_EXTENSION
    )


def get_files_from_folder(folder_path, only_files_with_extension=""):
    """
    Function to get file paths to files in directory

    Parameters
    ----------
    folder_path : str
        directory to extract files from
    only_files_with_extension : str, optional
        if provided, will only extract
            files that end with this extension, by default ""

    Returns
    -------
    file_paths : str iterable
        absolute file paths
    """

    return [
        join(folder_path, file)
        for file in listdir(folder_path)
        if file.endswith(only_files_with_extension)
    ]


def get_subfolders(folder_path):
    """
    Function to get subfolders in directory

    Parameters
    ----------
    folder_path : str
        directory to extract files from

    Returns
    -------
    subfolder_paths : str iterable
        absolute folder paths to subfolders
    """

    return [
        join(folder_path, folder)
        for folder in listdir(folder_path)
        if isdir(join(folder_path, folder))
    ]


def get_filename(file_path):
    """
    Function to convert a file's path to the filename

    Parameters
    ----------
    file_path : str
        file path to be parsed

    Returns
    -------
    filename : str
        filename without extension
    """

    if ("\\" in file_path) or ("/" in file_path):
        # strip path to file
        file_path = split(file_path)[1]
    return file_path.split(".")[0]  # remove extension if any


def get_path_to_file(file_path):
    """
    Function to get absolute path to folder containing file

    Parameters
    ----------
    file_path : str
        absolute path to file

    Returns
    -------
    folder_path : str
        absolute path to folder containing file
    """

    return split(file_path)[0]


def create_folder_if_not_exists(folder_path):
    """
    Function to create a folder, and all intermediate directories,
    if they don't exist

    Parameters
    ----------
    folder_path : str
        absolute folder path

    Returns
    -------
    folder_path : str
        absolute folder path
    """

    if not exists(folder_path):
        makedirs(folder_path)
    return folder_path


def generate_video_from_images(
    folder_path, video_name, figsize=None, fps=4, delete_pictures_when_done=False
):
    """
    Function to combine a series of images into an .mp4 video

    Parameters
    ----------
    folder_path : str
        absolute folder path to folder containing images
    video_name : str
        filename the final video will be saved with
    figsize : tuple, optional
        dimensions of the video frames i.e. (width:int, height:int), by default None.
        If not provided, the dimensions of the first image in the folder will be used
    fps : int, optional
        number of frames/images per second shown in the video, by default 4
    delete_pictures_when_done : bool, optional
        whether to delete all images in the folder_path
        after generating the video, by default False

    Notes
    -----
    images in the folder_path should be ordered in the order they appear
    in the video. This can be done by naming the video using indexes
    e.g. image_001, image_002
    All files in folder_path must be images
    """

    if os.path.exists(video_name):
        os.remove(video_name)

    image_paths = get_files_from_folder(folder_path)
    if image_paths:
        if figsize is None:
            image = cv2.imread(os.path.join(folder_path, image_paths[0]))
            figsize = (image.shape[1], image.shape[0])

        # set video parameters
        fc = cv2.VideoWriter_fourcc(*"mp4v")
        video = cv2.VideoWriter(video_name, fc, fps, figsize)

        # combine the images into the video
        for image_path in image_paths:
            try:
                image = cv2.imread(os.path.join(folder_path, image_path))
                image = cv2.resize(image, figsize)
                video.write(image)
                if delete_pictures_when_done:
                    os.remove(os.path.join(folder_path, image_path))
            except:
                pass
        video.release()
