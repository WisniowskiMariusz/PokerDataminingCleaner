# PokerDataminingCleaner
This is a Python script which is erasing data about chosen players from datamining poker packages.

Problem:
When player is playing a hand on his side this hands is already in his database. When we give him package with hand history of the same hand then it will be duplicate in his database.

Solution:
We are searching whole package for players IDs which client give to us (strictly confidental) and erasing all hands played buy player (client).

Result:
Clients get all datamined hand history but not this which he already have in his database (played by himself).
