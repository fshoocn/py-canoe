from typing import Union

import win32com.client


class DatabaseSetup:
    def __init__(self, com_object) -> None:
        self.com_object = com_object

    @property
    def databases(self) -> 'Databases':
        return Databases(self.com_object.Databases)

class Database:
    """The Database object represents the assigned database of the CANoe application."""
    def __init__(self, com_object):
        self.com_object = win32com.client.Dispatch(com_object)

    @property
    def channel(self) -> int:
        return self.com_object.Channel

    @channel.setter
    def channel(self, channel: int) -> None:
        self.com_object.Channel = channel

    @property
    def full_name(self) -> str:
        return self.com_object.FullName

    @full_name.setter
    def full_name(self, full_name: str) -> None:
        self.com_object.FullName = full_name

    @property
    def name(self) -> str:
        return self.com_object.Name

    @property
    def path(self) -> str:
        return self.com_object.Path


class Databases:
    """The Databases object represents the assigned databases of CANoe."""
    def __init__(self, com_object):
        self.com_object = com_object

    @property
    def count(self) -> int:
        return self.com_object.Count
    
    def item(self,index: int =None) -> Union[list[Database], Database]:
        """返回所有数据库对象"""
        if index is None:
            return [
                Database(self.com_object.Item(index)) 
                for index in range(1, self.com_object.Count + 1)
            ]
        else:
            return Database(self.com_object.Item(index))

    def add(self, full_name: str) -> 'Database':
        return Database(self.com_object.Add(full_name))

    def add_network(self, database_name: str, network_name: str) -> 'Database':
        return Database(self.com_object.AddNetwork(database_name, network_name))

    def remove(self, index: int) -> None:
        self.com_object.Remove(index)

