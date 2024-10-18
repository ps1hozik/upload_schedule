from pymongo import MongoClient
import os


def get_database():
    client = MongoClient(os.getenv("CONNECTION_STRING"))
    return client["schedule_bot"]


def delete_collections_except_users():
    db = get_database()
    collections = db.list_collection_names()
    for collection in collections:
        if collection != "new_users":
            db[collection].drop()
            print(f"Collection '{collection}' has been dropped.")


if __name__ == "__main__":
    delete_collections_except_users()
    # dbname = get_database()
