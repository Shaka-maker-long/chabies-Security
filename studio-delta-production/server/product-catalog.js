"use strict";

const DIMENSIONS = [
  ["Amalia Side Table", 560, 550, 550, ""],
  ["Amara Bed Double", 1200, 1400, 1900, ""],
  ["Angelica Coffee Table", 475, 474, "", 850],
  ["Blaire Bedside Cabinet", 600, 400, 400, ""],
  ["Blaire Dresser", 1385, 700, 430, ""],
  ["Blaire Curved Headboard", 1150, 2520, 25, ""],
  ["Thandi Bedside Cabinet", 600, 600, 400, ""],
  ["Olivia Bathroom Cabinet", 700, 460, 120, ""],
  ["Madeleine Room Divider", 1800, 520, 20, ""],
  ["Melody Standing Mirror", 1760, 500, 400, ""],
  ["Amelia Arched Cabinet", 2100, 1000, 450, ""],
  ["Nandi Dining Set", 800, 500, 1000, ""],
  ["Camilla Display Cabinet", 1700, 820, 370, ""],
  ["Delilah Display Cabinet", 2100, 350, 450, ""],
  ["Khanya Kitchen Cabinet", 1100, 2320, 320, ""],
  ["Luna Console Table", 820, 1200, 370, ""],
  ["Serena Sideboard", 820, 1670, 450, ""],
  ["Thalia Display Cabinet", 1800, 825, 440, ""],
  ["Thandi Display Cabinet", 1800, 820, 430, ""],
  ["Angelica Patio Chair", 777, 750, 820, ""],
  ["Angelica Patio Couch", 777, 2100, 830, ""],
  ["Angelica Planter Box (Large)", 800, 350, 350, ""],
  ["Blaire Dressing Table", 790, 1200, 480, ""],
  ["Blaire Platform Bed", 300, 2020, 1870, ""],
  ["Blaire Wardrobe", 1800, 1200, 600, ""],
  ["Brooke Open Wardrobe", 1800, 1200, 440, ""],
  ["Charlene Arched Cabinet", 2225, 980, 350, ""],
  ["Charlotte Chair", 850, 710, 550, ""],
  ["Charmaine Overhead Shelf", 800, 1640, 300, ""],
  ["Daphne Rectangular Mirror", 1750, 1050, 25, ""],
  ["Diana Cabinet", 910, 840, 450, ""],
  ["Ella Arched Cabinet", 2130, 470, 400, ""],
  ["Eloise Desk", 750, 1200, 640, ""],
  ["Eve Patio Bench", 440, 1885, 410, ""],
  ["Eve Patio Set", 700, 2040, 790, ""],
  ["Eve Patio Table", 700, 2040, 790, ""],
  ["Evelyn Cabinet", 1800, 820, 350, ""],
  ["Evie Arched Wardrobe", 1700, 900, 450, ""],
  ["Evie Changing Compactum", 975, 1240, 500, ""],
  ["Evie Crib Cot", 950, 1400, 700, ""],
  ["Evie Toy Shelf", 650, 1300, 400, ""],
  ["Fleur TV Unit", 600, 1670, 450, ""],
  ["Giselle Chair", 860, 700, 850, ""],
  ["Imani Open Shelf", 2100, 970, 430, ""],
  ["Jasmine Planter", 1000, 300, 300, ""],
  ["Leonora Bedside Table", 600, 600, 400, ""],
  ["Lila Side Table", 500, "", "", 500],
  ["Lindi Console Table", 820, 1200, 370, ""],
  ["Lorna Open Shelf", 1800, 540, 378, ""],
  ["Lucia Coffee Table", 490, 900, 500, ""],
  ["Margaux Open Wardrobe", 1800, 1200, 440, ""],
  ["Maria Desk", 750, 1500, 800, ""],
  ["Naledi Drinks Cabinet", 1600, 700, 300, ""],
  ["Nandi Dining Chair", 725, 445, 495, ""],
  ["Nandi Dining Table", 800, 500, 1000, ""],
  ["Naomi Arched Cabinet", 2100, 1000, 450, ""],
  ["Noelle Open Shelf", 1800, 1040, 338, ""],
  ["Nova Bar Stool", 885, 350, 350, ""],
  ["Ophelia Coffee Table", 375, 850, 850, ""],
  ["Penelope Display Cabinet", 1940, 820, 460, ""],
  ["Raye Coffee Table", 375, 1070, 595, ""],
  ["Rosie Planter Medium", 1000, 350, 350, ""],
  ["Rosie Planter Small", 700, 350, 350, ""],
  ["Rosie Planter Tall", 1300, 350, 350, ""],
  ["Sienna Buffet", 900, 2500, 500, ""],
  ["Tumi TV Unit", 360, 1600, 350, ""],
  ["Uriah Side Table", 690, 300, 400, ""],
  ["Valerie Drinks Cabinet", 1700, 1000, 400, ""],
  ["Violet Sideboard 3 - Door", 820, 1260, 450, ""],
  ["Violet Sideboard 4 - Door", 820, 1670, 450, ""],
  ["Vivienne Arched Cabinet", 2225, 980, 350, ""],
  ["Zahara Arched Mirror", 1800, 600, 25, ""],
  ["Zola Console Table", 800, 500, 1475, ""],
  ["Zola Side Table", 500, 465, 440, ""],
  ["Aliana Dining Table", 775, "", "", 1200],
  ["Zuri Oval Sideboard", 800, 1300, 500, ""],
  ["Palesa Coffee Table", 450, 1600, 800, ""],
  ["Felicity Open Shelf Large", 300, 900, 250, ""],
  ["Felicity Open Shelf Medium", 300, 600, 250, ""],
  ["Felicity Open Shelf Small", 300, 450, 250, ""],
  ["Felicity Open Shelf Bundle", 300, 900, 250, ""],
  ["Hex Bar Stool", 837, 462, 622, ""],
  ["New Design", 1, 1, 1, ""],
  ["Talitha Bookshelf", 1460, 560, 560, ""],
  ["Amara Bed", 1200, 1100, 1900, ""],
  ["Oakland Sideboard 4-Door", 800, 2200, 500, ""],
  ["Estelle Patio Table", 750, 1600, 700, ""],
  ["Violet Sideboard 2 - Door", 820, 840, 450, ""],
  ["Tatiana Bookshelf", 1400, 250, 250, ""],
  ["Mila Dining Table", 733, 600, 600, 600],
  ["Zandile TV Unit", 460, 1400, 400, ""],
  ["Nola Side Table", 500, 487, 487, 475],
  ["Marisol Coffee Table", 490, 1260, 435, ""],
  ["Kyra Nesting Side Tables", 600, 450, 450, ""],
  ["Vivienne Sideboard", 950, 1940, 500, ""],
  ["Brooklyn Display Cabinet", 2315, 500, 405, ""],
  ["Michelle Dining Table", 750, 1540, 790, ""],
  ["Vanessa Floating Shelf Medium", 300, 600, 250, ""],
  ["Air Chair", 835, 445, 580, ""],
  ["Anaya Display Cabinet", 1800, 825, 440, ""],
  ["Ruby Drinks Cabinet", 1700, 820, 500, ""],
  ["Brie Drinks Cabinet", 1800, 820, 430, ""],
  ["Astra Steel Cabinet", 1800, 840, 450, ""],
  ["Brielle Bedside Cabinet", 600, 500, 350, ""],
  ["Custom Steel Frame", 1480, 650, 650, ""],
  ["Zanele TV Unit", 600, 2200, 390, ""],
  ["Oakland Room Divider", 2100, 1500, 485, ""],
  ["Vinette Wine Cabinet", 1800, 820, 430, ""],
  ["Michelle Coffee Table", 470, 840, 840, ""],
  ["Claire Coffee Table", 470, 840, 840, ""],
];

const IMAGES = {
  "Amalia Side Table": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO-DELTA_Round-Steel-Tables_View-C.jpg",
  "Amara Bed Double": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO-DELTA_INDUSTRIAL-BED_VIEW-A.jpg",
  "Angelica Coffee Table": "https://www.studiodelta.co.za/wp-content/uploads/2024/04/STUDIO-DELTA_MODERN-TUSCAN-COFFEE-TABLE_VIEW-E-600x600.webp",
  "Blaire Bedside Cabinet": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO-DELTA_SINGLE-DOOR-BEDSIDE-CABINET_VIEW-B-600x600.jpg.webp",
  "Blaire Dresser": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO-DELTA_TALLBOY-3-DRAWER-DRESSER_VIEW-D-600x600.jpg.webp",
  "Blaire Curved Headboard": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO-DELTA_CURVED-STEEL-UPHOLSTERED-HEADBOARD_VIEW-F-1.jpg",
  "Thandi Bedside Cabinet": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO-DELTA_REEDED-BEDSIDE-TABLE_VIEW-B.jpg",
  "Olivia Bathroom Cabinet": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO-DELTA_Mirror-Bathroom-Cabinet_viewA-600x600.jpg.webp",
  "Madeleine Room Divider": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO-DELTA_REEDED-GLASS-ROOM-DIVIDER_VIEW-C-1.jpg",
  "Melody Standing Mirror": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO-DELTA_WARDROBE-MIRROR_VIEW-A.jpg",
  "Amelia Arched Cabinet": "https://www.studiodelta.co.za/wp-content/uploads/2024/06/STUDIO-DELTA_ARCHED-STORAGE-CABINET-IVORY-AND-CLEAR-GLASS_VIEW-B-600x600.webp",
  "Nandi Dining Set": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO-DELTA_SPACE-SAVER-DINING-CHAIR_VIEW-A.jpg",
  "Angelica Patio Chair": "https://www.studiodelta.co.za/wp-content/uploads/2024/04/STUDIO-DELTA_MODERN-TUSCAN-SINGLE-SEATER_VIEW-B-600x600.webp",
  "Angelica Patio Couch": "https://www.studiodelta.co.za/wp-content/uploads/2024/04/STUDIO-DELTA_MODERN-TUSCAN-THREE-SEATER_VIEW-B-600x600.webp",
  "Angelica Planter Box (Large)": "https://www.studiodelta.co.za/wp-content/uploads/2024/04/STUDIO-DELTA_MODERN-TUSCAN-PLANTERS-VIEW-AA-600x600.webp",
  "Blaire Dressing Table": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO-DELTA_BEDROOM-COLLECTION_Steel-Dressing-Table-Vanity-with-Wooden-Drawers_view-A.jpg",
  "Blaire Platform Bed": "https://www.studiodelta.co.za/wp-content/uploads/2023/11/STUDIO-DELTA_FLOATING-BEDFRAME_viewB-600x600.jpg.webp",
  "Blaire Wardrobe": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO-DELTA_REEDED-GLASS-WARDROBE-WITH-WOODEN-DRAWERS_VIEW-A-1-600x600.jpg.webp",
  "Brooke Open Wardrobe": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO-DELTA_BROOKE-OPEN-WARDROBE_View-F-600x600.webp",
  "Camilla Display Cabinet": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO-DELTA_RETRO-DISPLAY-CABINET_VIEW-D-BLACK.jpg",
  "Charlene Arched Cabinet": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/Studio-Delta_Arched-Glass-Cabinet_viewA-1-600x600.jpg.webp",
  "Charlotte Chair": "https://www.studiodelta.co.za/wp-content/uploads/2023/11/STUDIO-DELTA_UPHOLSTERED-STEEL-ROUNDED-CHAIRS_VIEW-A-600x600.jpg.webp",
  "Charmaine Overhead Shelf": "https://www.studiodelta.co.za/wp-content/uploads/2025/01/STUDIO-DELTA_CHARMAINE-OVERHEAD-SHELF_VIEW-E-600x600.webp",
  "Daphne Rectangular Mirror": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO-DELTA_STANDING-MIRROR_VIEW-B-600x600.jpg.webp",
  "Delilah Display Cabinet": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO-DELTA_NARROW-GLASS-DISPLAY-CABINET-VIEW-B.jpg",
  "Diana Cabinet": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO-DETLA_SHORT-STEEL-CABINET-WITH-PERFORATED-DOORS_VIEW-A-600x600.jpg.webp",
  "Ella Arched Cabinet": "https://www.studiodelta.co.za/wp-content/uploads/2024/07/STUDIO-DELTA_ARCHED-STORAGE-CABINET-Narrow_VIEW-B-600x600.webp",
  "Eloise Desk": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO-DELTA_MINIMALIST-DESK_VIEW-B-1.jpg",
  "Eve Patio Bench": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO-DELTA_INDUSTRIAL-PATIO-BENCH_VIEW-A-600x600.webp",
  "Eve Patio Set": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO-DELTA_INDUSTRIAL-PATIO-TABLE-AND-2-BENCHES_VIEW-A-600x600.webp",
  "Eve Patio Table": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO-DELTA_INDUSTRIAL-STEEL-AND-WOOD-TABLE_VIEW-B-600x600.jpg.webp",
  "Evelyn Cabinet": "https://www.studiodelta.co.za/wp-content/uploads/2024/06/STUDIO-DELTA_MODERN-ANTIQUE-GLASS-CABINET_VIEW-C-600x600.webp",
  "Evie Arched Wardrobe": "https://www.studiodelta.co.za/wp-content/uploads/2023/12/STUDIO-DELTA_ARCHED-UPHOLSTERED-CABINET_VIEW-A-600x600.webp",
  "Evie Changing Compactum": "https://www.studiodelta.co.za/wp-content/uploads/2023/11/STUDIO-DELTA_BABY-CHANGING-COMPACTUM-_VIEW-A-600x600.jpg.webp",
  "Evie Crib Cot": "https://www.studiodelta.co.za/wp-content/uploads/2023/12/STUDIO-DELTA_ADJUSTABLE-UPHOLSTERED-BABY-CRIB-COT_VIEW-A-1-600x600.jpg.webp",
  "Evie Toy Shelf": "https://www.studiodelta.co.za/wp-content/uploads/2023/11/STUDIO-DELTA_TOY-STORAGE-SHELF_VIEW-A-600x600.jpg.webp",
  "Fleur TV Unit": "https://www.studiodelta.co.za/wp-content/uploads/2024/07/STUDIO-DELTA_RETRO-TV-CABINET_VIEW-A-600x600.webp",
  "Giselle Chair": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO-DELTA_MINIMALIST-STEEL-AND-WOOD-CHAIR_VIEW-A-600x600.jpg.webp",
  "Imani Open Shelf": "https://www.studiodelta.co.za/wp-content/uploads/2024/04/STUDIO-DELTA_MINIMALIST-ARCHED-OPEN-SHELF_VIEW-F.jpg",
  "Jasmine Planter": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO-DELTA_STEEL-PLANTERS_VIEW-BB.webp",
  "Khanya Kitchen Cabinet": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/Studio-Delta_GLASS-STEEL-KITCHEN-CABINET_viewB-1.jpg",
  "Leonora Bedside Table": "https://www.studiodelta.co.za/wp-content/uploads/2024/09/STUDIO-DELTA_SEAMLESS-STEEL-AND-WOOD-BEDSIDE-TABLE_VIEW-B.webp",
  "Lila Side Table": "https://www.studiodelta.co.za/wp-content/uploads/2025/01/STUDIO-DELTA_LILA-SIDE-TABLE_VIEW-B-600x600.webp",
  "Lindi Console Table": "https://www.studiodelta.co.za/wp-content/uploads/2024/07/STUDIO-DELTA_CURVED-STEEL-AND-WOOD-CONSOLE_VIEW-A-600x600.webp",
  "Lorna Open Shelf": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO-DELTA_NARROW-STEEL-AND-GLASS-OPEN-SHELF_VIEW-A-2-600x750.jpg.webp",
  "Lucia Coffee Table": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO-DELTA_NESTED-COFFEE-TABLES_VIEW-B-600x600.jpg.webp",
  "Luna Console Table": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO-DELTA_STEEL-AND-GLASS-CONSOLE-TABLE_VIEW-A-600x600.jpg.webp",
  "Margaux Open Wardrobe": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO-DELTA_INDUSTRIAL-OPEN-WARDROBE_VIEW-B-600x664.jpg.webp",
  "Maria Desk": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO-DELTA_STEEL-AND-WOOD-OFFICE-DESK_VIEW-B-600x600.jpg.webp",
  "Naledi Drinks Cabinet": "https://www.studiodelta.co.za/wp-content/uploads/2024/09/STUDIO-DELTA_PILL-SHAPED-DRINKS-CABINET_VIEW-B.jpg",
  "Nandi Dining Chair": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO-DELTA_SPACE-SAVER-DINING-CHAIR_VIEW-D-600x600.jpg.webp",
  "Nandi Dining Table": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO-DELTA_SPACE-SAVER-DINING-CHAIR_VIEW-A-600x600.jpg.webp",
  "Naomi Arched Cabinet": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO-DELTA_ARCHED-STORAGE-CABINET_VIEW-B1.jpg",
  "Noelle Open Shelf": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO_DELTA_STEEL-AND-GLASS-OPEN-SHELF_VIEW-B.jpg",
  "Nova Bar Stool": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO-DELTA_KITCHEN-BAR-STOOL_VIEW-A-600x600.jpg.webp",
  "Ophelia Coffee Table": "https://www.studiodelta.co.za/wp-content/uploads/2024/07/STUDIO-DELTA_SQUARE-MINIMALIST-STEEL-AND-GLASS-COFFEE-TABLE_VIEW-C.jpg",
  "Penelope Display Cabinet": "https://www.studiodelta.co.za/wp-content/uploads/2024/06/MINIMALIST-REEDED-GLASS-DISPLAY-CABINET_-VIEW-A-600x600.webp",
  "Raye Coffee Table": "https://www.studiodelta.co.za/wp-content/uploads/2024/08/STUDIO-DELTA_RECTANGULAR-GLASS-AND-STEEL-COFFEE-TABLE_VIEW-D-600x600.webp",
  "Rosie Planter Medium": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO-DELTA_Round-steel-planters_view-E-scaled-1.jpg",
  "Rosie Planter Small": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO-DELTA_Round-steel-planters_view-E-scaled-1.jpg",
  "Rosie Planter Tall": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO-DELTA_Round-steel-planters_view-E-scaled-1.jpg",
  "Serena Sideboard": "https://www.studiodelta.co.za/wp-content/uploads/2024/06/STUDIO-DELTA_RETRO-SIDEBOARD-WITH-OCEAN-VUE-GLASS_VIEW-A-600x600.webp",
  "Sienna Buffet": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO-DELTA_LARGE-INDUSTRIAL-BUFFET_VIEW-D-600x600.webp",
  "Thalia Display Cabinet": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO-DELTA_CLEAR-GLASS-DISPLAY-CABINET_VIEW-A-1.jpg",
  "Thandi Display Cabinet": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO-DELTA_REEDED-GLASS-DISPLAY-CABINET_VIEW-A.jpg",
  "Tumi TV Unit": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO-DELTA_MINIMAL-STEEL-TV-UNIT_VIEW-D-600x600.jpg.webp",
  "Uriah Side Table": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO-DELTA_U-SIDE-TABLE_VIEW-C-600x600.jpg.webp",
  "Valerie Drinks Cabinet": "https://www.studiodelta.co.za/wp-content/uploads/2024/06/INDUSTRIAL-STEEL-DRINKS-CABINET_VIEW-A-600x600.webp",
  "Violet Sideboard 3 - Door": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO-DELTA_INDUSTRIAL-STEEL-SIDEBOARD-3-DOOR_VIEW-A-600x600.jpg.webp",
  "Violet Sideboard 4 - Door": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO-DELTA_INDUSTRIAL-STEEL-SIDEBOARD_4-DOOR_VIEW-D-1.jpg",
  "Vivienne Arched Cabinet": "https://www.studiodelta.co.za/wp-content/uploads/2024/09/STUDIO-DELTA_DUAL-TONED-ARCHED-GLASS-DISPLAY-CABINET_VIEW-E-600x600.jpg.webp",
  "Zahara Arched Mirror": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO-DELTA_ARCHED-STANDING-MIRROR_VIEW-D-600x600.jpg.webp",
  "Zola Console Table": "https://www.studiodelta.co.za/wp-content/uploads/2024/07/STUDIO-DELTA_STEEL-AND-GLASS-SERVER_VIEW-B-600x600.webp",
  "Zola Side Table": "https://www.studiodelta.co.za/wp-content/uploads/2024/07/STUDIO-DELTA_ABSTRACT-GLASS-AND-STEEL-SIDE-TABLE_VIEW-A-600x600.webp",
  "Aliana Dining Table": "https://www.studiodelta.co.za/wp-content/uploads/2024/09/STUDIO-DELTA_ROUND-STEEL-AND-GLASS-DINING-TABLE_VIEW-C-600x600.webp",
  "Zuri Oval Sideboard": "https://www.studiodelta.co.za/wp-content/uploads/2025/01/STUDIO-DELTA_ZURI-SIDEBOARD_VIEW-D.webp",
  "Palesa Coffee Table": "https://www.studiodelta.co.za/wp-content/uploads/2025/04/STUDIO-DELTA_PALESA-COFFEE-TABLE_VIEW-A.webp",
  "Felicity Floating Shelf": "https://www.studiodelta.co.za/wp-content/uploads/2025/01/FELICITY-FLOATING-SHELF_VIEW-B.webp",
  "Felicity Open Shelf Large": "https://www.studiodelta.co.za/wp-content/uploads/2025/01/FELICITY-FLOATING-SHELF_VIEW-B.webp",
  "Felicity Open Shelf Medium": "https://www.studiodelta.co.za/wp-content/uploads/2025/01/FELICITY-FLOATING-SHELF_VIEW-B.webp",
  "Felicity Open Shelf Small": "https://www.studiodelta.co.za/wp-content/uploads/2025/01/FELICITY-FLOATING-SHELF_VIEW-B.webp",
  "Felicity Open Shelf Bundle": "https://www.studiodelta.co.za/wp-content/uploads/2025/01/FELICITY-FLOATING-SHELF_VIEW-B.webp",
  "Talitha Bookshelf": "https://www.studiodelta.co.za/wp-content/uploads/2025/02/STUDIO-DELTA_TALITHA-BOOKSHELF_VIEW-B.webp",
  "Amara Bed": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO-DELTA_INDUSTRIAL-BED_VIEW-A-600x600.jpg.webp",
  "Oakland Sideboard 4-Door": "https://www.studiodelta.co.za/wp-content/uploads/2025/06/STUDIO-DELTA_OAKLAND-4-DOOR-SIDEBOARD_View-A.webp",
  "Estelle Patio Table": "https://www.studiodelta.co.za/wp-content/uploads/2023/11/STUDIO-DELTA_MINIMALIST-STEEL-AND-WOOD-PATIO-TABLE_VIEW-A.jpg",
  "Violet Sideboard 2 - Door": "https://www.studiodelta.co.za/wp-content/uploads/2023/07/STUDIO-DELTA_INDUSTRIAL-STEEL-SIDEBOARD-2-DOOR_VIEW-C1-600x600.jpg.webp",
  "Tatiana Bookshelf": "https://www.studiodelta.co.za/wp-content/uploads/2025/02/STUDIO-DELTA_TATIANA-BOOKSHELF_VIEW-B.webp",
  "Mila Dining Table": "https://www.studiodelta.co.za/wp-content/uploads/2025/08/STUDIO-DELTA-MILA-DINING-TABLE-C.webp",
  "Zandile TV Unit": "https://www.studiodelta.co.za/wp-content/uploads/2025/05/STUDIO-DELTA_ZANDILE-TV-UNIT_VIEW-F.webp",
  "Nola Side Table": "https://www.studiodelta.co.za/wp-content/uploads/2025/01/STUDIO-DELTA_NOLA-SIDE-TABLE_VIEW-A.webp",
  "Marisol Coffee Table": "https://www.studiodelta.co.za/wp-content/uploads/2025/04/STUDIO-DELTA_MARISOL-COFFEE-TABLE_VIEW-A.webp",
  "Kyra Nesting Side Tables": "https://www.studiodelta.co.za/wp-content/uploads/2025/04/STUDIO-DELTA_KYRA-NESTED-SIDE-TABLES_VIEW-B.webp",
  "Vivienne Sideboard": "https://www.studiodelta.co.za/wp-content/uploads/2025/05/STUDIO-DELTA_VIVIENNE-SIDEBOARD-4DOOR_VIEW-A.webp",
  "Brooklyn Display Cabinet": "https://www.studiodelta.co.za/wp-content/uploads/2025/07/STUDIO-DELTA_BROOKLYN-DISPLAY-CABINET_C.webp",
  "Michelle Dining Table": "https://www.studiodelta.co.za/wp-content/uploads/2025/08/STUDIO-DELTA-MICHELLE-DINING-TABLE-A.webp",
  "Vanessa Floating Shelf Medium": "https://www.studiodelta.co.za/wp-content/uploads/2025/01/STUDIO-DELTA_VANESSA-FLOATING-SHELF-MEDIUM_VIEW-B.webp",
  "Anaya Display Cabinet": "https://www.studiodelta.co.za/wp-content/uploads/2025/10/ANAYA-DISPLAY-CABINET_Studio-Delta_1.webp",
  "Ruby Drinks Cabinet": "https://www.studiodelta.co.za/wp-content/uploads/2025/09/Studio-Delta-Ruby-Drinks-Cabinet_2.webp",
  "Zanele TV Unit": "https://www.studiodelta.co.za/wp-content/uploads/2025/05/STUDIO-DELTA_ZANELE-TV-UNIT_VIEW-C.webp",
  "Oakland Room Divider": "https://www.studiodelta.co.za/wp-content/uploads/2025/07/STUDIO-DELTA-OAKLAND-ROOM-DIVIDER_2.webp",
};

const ALIASES = {
  "felicity floating shelf large": "Felicity Open Shelf Large",
  "felicity floating shelf medium": "Felicity Open Shelf Medium",
  "felicity floating shelf small": "Felicity Open Shelf Small",
  "felicity floating shelf bundle": "Felicity Open Shelf Bundle",
  "felicity floating shelf": "Felicity Open Shelf Large",
  "oakland sideboard 4 door": "Oakland Sideboard 4-Door",
  "violet sideboard 3 door": "Violet Sideboard 3 - Door",
  "violet sideboard 4 door": "Violet Sideboard 4 - Door",
  "violet sideboard 2 door": "Violet Sideboard 2 - Door",
};

function normalizeName(name) {
  return String(name || "")
    .toLowerCase()
    .replace(/&/g, " and ")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function dimRecord(row) {
  return {
    name: row[0],
    height: row[1] === "" ? null : Number(row[1]),
    width: row[2] === "" ? null : Number(row[2]),
    depth: row[3] === "" ? null : Number(row[3]),
    diameter: row[4] === "" ? null : Number(row[4]),
  };
}

const BY_KEY = new Map();
for (const row of DIMENSIONS) {
  const rec = dimRecord(row);
  rec.imageUrl = IMAGES[rec.name] || "";
  BY_KEY.set(normalizeName(rec.name), rec);
}

function lookupProduct(productName) {
  const key = normalizeName(productName);
  if (!key) return null;
  if (BY_KEY.has(key)) return { ...BY_KEY.get(key) };
  if (ALIASES[key]) {
    const aliased = BY_KEY.get(normalizeName(ALIASES[key]));
    if (aliased) return { ...aliased };
  }
  for (const [catalogKey, rec] of BY_KEY) {
    if (key.includes(catalogKey) || catalogKey.includes(key)) {
      return { ...rec };
    }
  }
  return null;
}

function listCatalog() {
  return DIMENSIONS.map((row) => {
    const rec = dimRecord(row);
    rec.imageUrl = IMAGES[rec.name] || "";
    return rec;
  });
}

function dimensionsForDisplay(productName, overrides) {
  const found = lookupProduct(productName) || {};
  const src = overrides && typeof overrides === "object" ? { ...found, ...overrides } : found;
  const rows = [];
  if (src.height != null && src.height !== "") rows.push({ name: "Height", value: src.height });
  if (src.width != null && src.width !== "") rows.push({ name: "Width", value: src.width });
  if (src.depth != null && src.depth !== "") rows.push({ name: "Depth", value: src.depth });
  if (src.diameter != null && src.diameter !== "") rows.push({ name: "Diameter", value: src.diameter });
  return rows;
}

function dimensionsString(productName, overrides) {
  const rows = dimensionsForDisplay(productName, overrides);
  if (!rows.length) return "Standard";
  return rows.map((r) => `${r.name}: ${r.value}mm`).join(", ");
}

module.exports = {
  normalizeName,
  lookupProduct,
  listCatalog,
  dimensionsForDisplay,
  dimensionsString,
  COMPANY_LOGO_URL: "https://studiodelta.co.za/wp-content/uploads/2024/03/Studio-Delta_company_logo.jpg",
};
