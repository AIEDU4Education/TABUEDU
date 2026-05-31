export type CardStatus = 'Inicio' | 'Proceso' | 'Logrado' | 'Destacado';

export interface CardData {
  id: string; // Proporciona un ID único para operaciones CRUD
  secreto: string; // Palabra principal a adivinar (también referida como word)
  tabu: string[]; // Palabras prohibidas (también referida como tabooWords)
  nivelLogro: CardStatus | string; // 'Inicio' | 'Proceso' | 'Logrado' | 'Destacado'
  categoria: string; // Categoría o tema de la tarjeta
  pista?: string; // Pista de pie de tarjeta estilo socrático
  
  // Soporte de alias para máxima flexibilidad y cumplimiento estricto
  word?: string; // Alias de secreto
  tabooWords?: string[]; // Alias de tabu
  status?: CardStatus; // Alias de nivelLogro
}

export interface ExcelColumnMapping {
  wordColumn: string;
  tabooColumns: string[];
  tabooDelimiter: string;
  isSingleColumnForTaboo: boolean;
  statusColumn?: string;
  pistaColumn?: string;
  categoriaColumn?: string;
}
