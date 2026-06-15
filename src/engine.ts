import type { GroupState, AnswerOutcome } from './data/gameContent.types';

export const GROUP_COLORS: Record<string, string> = {
  blue:   '#1E6ABF',
  green:  '#1A7A3A',
  orange: '#C05A10',
  purple: '#7A3AAA',
};

export const GROUP_BG: Record<string, string> = {
  blue:   '#DDEAFF',
  green:  '#D8F0E0',
  orange: '#FFE8CC',
  purple: '#EDD8FF',
};

export function initGroupState(groupId: string): GroupState {
  return {
    groupId,
    currentStationOrder: 6, // pilot: all groups start at מדד תפוקה
    points: 0,
    errors: 0,
    waitRoundsRemaining: 0,
    pendingAutoAdvance: false,
    usedUnexpectedEventIds: [],
    completedStationIds: [],
    playedThisRound: false,
  };
}

export function applyOutcome(
  state: GroupState,
  outcome: AnswerOutcome,
  maxOrder: number,
  stationId: string,
): GroupState {
  const completed = state.completedStationIds.includes(stationId)
    ? state.completedStationIds
    : [...state.completedStationIds, stationId];

  if (outcome === 'correct') {
    return {
      ...state,
      currentStationOrder: Math.min(state.currentStationOrder + 1, maxOrder),
      points: state.points + 10,
      pendingAutoAdvance: false,
      completedStationIds: completed,
      playedThisRound: true,
    };
  }
  if (outcome === 'partial') {
    return {
      ...state,
      points: state.points + 5,
      errors: state.errors + 1,
      pendingAutoAdvance: true,
      completedStationIds: completed,
      playedThisRound: true,
    };
  }
  // wrong — go back 1, can't go below order 0
  const newOrder = state.currentStationOrder <= 0 ? 0 : state.currentStationOrder - 1;
  return {
    ...state,
    errors: state.errors + 1,
    currentStationOrder: newOrder,
    pendingAutoAdvance: false,
    playedThisRound: true,
  };
}

export function startNewRound(states: Record<string, GroupState>): Record<string, GroupState> {
  const next: Record<string, GroupState> = {};
  for (const [id, s] of Object.entries(states)) {
    let ns = { ...s, playedThisRound: false };
    if (ns.waitRoundsRemaining > 0) ns.waitRoundsRemaining -= 1;
    if (ns.pendingAutoAdvance && ns.waitRoundsRemaining === 0) {
      ns.currentStationOrder += 1;
      ns.pendingAutoAdvance = false;
    }
    next[id] = ns;
  }
  return next;
}
