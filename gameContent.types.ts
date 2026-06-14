export type StationType = 'professional' | 'unexpected_event' | 'transition' | 'reflection';
export type AnswerOutcome = 'correct' | 'partial' | 'wrong';

export interface Group {
  id: string; label: string; divisionName: string; colorToken: string;
  mobileFolderPath?: string; qrCodeUrl?: string;
}
export interface Station {
  id: string; name: string; month?: string; type: StationType; order: number;
  segment: string; boardVisibility: string; facilitatorVisibility: boolean;
  mobileRelevance: boolean; questionId?: string; unexpectedEventSetId?: string;
}
export interface GameState {
  groupId: string; currentStationOrder: number; points: number;
  waitingOneRound: boolean; playedThisRound: boolean;
  usedUnexpectedEventIds: string[]; completedStationIds: string[];
}
