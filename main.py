import time

from download import download
from upload import upload


def main():
    download()
    time.sleep(5)
    errors = upload()
    print(f"Errors: {errors}")
    input("Press enter to exit\n")


if __name__ == "__main__":
    main()
