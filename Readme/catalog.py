"""Controlled vocabularies and the alias maps used by the cleaning stage."""

RAW_COLUMNS = [
    "OrderID",
    "OrderDate",
    "Region",
    "Channel",
    "Category",
    "Product",
    "Segment",
    "Quantity",
    "UnitPrice",
    "Discount",
    "UnitCost",
    "ShipDays",
    "Status",
    "SalesRep",
]

REGIONS = ["North", "South", "East", "West"]
CHANNELS = ["Store", "Online", "Wholesale"]
CATEGORIES = ["Furniture", "Lighting", "Textiles", "Kitchen"]
STATUSES = ["Completed", "Returned", "Cancelled"]
SEGMENTS = ["Consumer", "Corporate", "Home Office"]
REPS = [
    "A. Shah",
    "M. Okonkwo",
    "L. Berger",
    "R. Ibarra",
    "S. Nguyen",
    "P. Rossi",
    "K. Mensah",
    "J. Patel",
]

# (product, base price, cost ratio)
CATALOG = {
    "Furniture": [
        ("Oak Desk", 840, 0.62),
        ("Linen Sofa", 1250, 0.66),
        ("Walnut Shelf", 220, 0.55),
        ("Lounge Chair", 480, 0.60),
    ],
    "Lighting": [
        ("Arc Lamp", 310, 0.48),
        ("Pendant Light", 180, 0.50),
        ("Brass Sconce", 140, 0.46),
    ],
    "Textiles": [
        ("Wool Throw", 75, 0.42),
        ("Cotton Duvet", 120, 0.44),
        ("Linen Curtain", 95, 0.45),
    ],
    "Kitchen": [
        ("Ceramic Set", 86, 0.38),
        ("Chef Knife", 64, 0.36),
        ("Cast Skillet", 54, 0.40),
    ],
}

REGION_ALIASES = {
    "north": "North",
    "nort": "North",
    "south": "South",
    "sout": "South",
    "east": "East",
    "eas": "East",
    "west": "West",
    "wes": "West",
}
CHANNEL_ALIASES = {
    "store": "Store",
    "retail": "Store",
    "in-store": "Store",
    "instore": "Store",
    "online": "Online",
    "web": "Online",
    "ecommerce": "Online",
    "e-commerce": "Online",
    "wholesale": "Wholesale",
    "b2b": "Wholesale",
}
CATEGORY_ALIASES = {name.lower(): name for name in CATEGORIES}
STATUS_ALIASES = {
    "completed": "Completed",
    "complete": "Completed",
    "compleated": "Completed",
    "closed": "Completed",
    "returned": "Returned",
    "return": "Returned",
    "retrned": "Returned",
    "cancelled": "Cancelled",
    "canceled": "Cancelled",
    "cancled": "Cancelled",
}
SEGMENT_ALIASES = {
    "consumer": "Consumer",
    "corporate": "Corporate",
    "home office": "Home Office",
    "homeoffice": "Home Office",
    "home-office": "Home Office",
}

REGION_DIRT = {
    "North": ["North", "NORTH", "north", "Nort", " North "],
    "South": ["South", "SOUTH", "south", "Sout"],
    "East": ["East", "east", "EAST", "Eas"],
    "West": ["West", "west", "WEST", " west "],
}
STATUS_DIRT = {
    "Completed": ["Completed", "complete", "Compleated", "COMPLETED"],
    "Returned": ["Returned", "RETURNED", "retrned"],
    "Cancelled": ["Cancelled", "cancled", "Canceled"],
}
CHANNEL_DIRT = {
    "Store": ["Store", "retail", "In-Store"],
    "Online": ["Online", "web", "e-commerce"],
    "Wholesale": ["Wholesale", "B2B"],
}
