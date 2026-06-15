export type StationType = 'professional' | 'unexpected_event' | 'transition' | 'reflection';
export type AnswerOutcome = 'correct' | 'partial' | 'wrong';
export type StationSegment = 'workplan' | 'transition_to_compensation' | 'compensation' | 'reflection';

export interface Group {
  id: string;
  label: string;
  divisionName: string;
  colorToken: string;
  mobileFolderPath?: string;
  qrCodeUrl?: string;
}

export interface BoardPosition { x: number; y: number; }

export interface Station {
  id: string;
  name: string;
  month?: string;
  type: StationType;
  order: number;
  segment: StationSegment;
  boardPosition: BoardPosition;
  questionId?: string;
  unexpectedEventSetId?: string;
  transitionMessageId?: string;
}

export interface AnswerByGroup {
  groupId: string;
  taskName: string;
  whatIsWrong: string;
  correctFix: string;
  facilitatorNotes: string | null;
}

export interface StationQuestion {
  id: string;
  stationId: string;
  questionText: string;
  professionalPrinciple: string;
  timerSeconds: number;
  answersByGroup: Record<string, AnswerByGroup>;
}

export interface VisibleTask {
  taskId: string;
  taskName: string;
  companyGoal: string;
  goal: string;
  outcomeMetric: string;
  classification: 'top10' | 'routine';
  taskDescription: string;
  annualAchievement: string;
  outputMetricType: string;
  milestones: { Q1: string; Q2: string; Q3: string; Q4: string };
}

export interface CompensationModel {
  modelName: string;
  rewardMetric: string;
  formula: string;
  weight: string;
  performanceThresholds: string;
  dataSource: string;
  notes?: string;
}

export interface UnexpectedEvent {
  resultId: string;
  label: string;
  description: string;
  waitRounds: number | string;
}

export interface TransitionMessage { title: string; body: string; }

// Runtime state
export interface GroupState {
  groupId: string;
  currentStationOrder: number;
  points: number;
  errors: number;
  waitRoundsRemaining: number;
  pendingAutoAdvance: boolean;
  usedUnexpectedEventIds: string[];
  completedStationIds: string[];
  playedThisRound: boolean;
}
