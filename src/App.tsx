import React, { useReducer } from 'react';
import data from './data/gameContent.json';
import type { GroupState, AnswerOutcome } from './data/gameContent.types';
import { initGroupState, applyOutcome, startNewRound } from './engine';
import FacilitatorView from './views/Facilitator';
import MobileView from './views/Mobile';

type View = 'facilitator' | `mobile_${string}`;

export interface AppState {
  view: View;
  groupStates: Record<string, GroupState>;
  openStationId: string | null;
  answerRevealed: boolean;
  currentTurnGroupIndex: number;
  roundNumber: number;
}

type Action =
  | { type: 'SET_VIEW'; view: View }
  | { type: 'OPEN_STATION'; stationId: string }
  | { type: 'CLOSE_STATION' }
  | { type: 'REVEAL_ANSWER' }
  | { type: 'RECORD_OUTCOME'; groupId: string; outcome: AnswerOutcome }
  | { type: 'NEXT_TURN' };

const maxOrder = Math.max(...data.stations.map(s => s.order));

function initState(): AppState {
  const groupStates: Record<string, GroupState> = {};
  data.groups.forEach(g => { groupStates[g.id] = initGroupState(g.id); });
  return {
    view: 'facilitator',
    groupStates,
    openStationId: null,
    answerRevealed: false,
    currentTurnGroupIndex: 0,
    roundNumber: 1,
  };
}

function reducer(state: AppState, action: Action): AppState {
  switch (action.type) {
    case 'SET_VIEW':
      return { ...state, view: action.view, openStationId: null, answerRevealed: false };
    case 'OPEN_STATION':
      return { ...state, openStationId: action.stationId, answerRevealed: false };
    case 'CLOSE_STATION':
      return { ...state, openStationId: null, answerRevealed: false };
    case 'REVEAL_ANSWER':
      return { ...state, answerRevealed: true };
    case 'RECORD_OUTCOME': {
      const gs = state.groupStates[action.groupId];
      const station = data.stations.find(s => s.id === state.openStationId);
      if (!gs || !station) return state;
      const updated = applyOutcome(gs, action.outcome, maxOrder, station.id);
      const newGroupStates = { ...state.groupStates, [action.groupId]: updated };
      const nextIndex = (state.currentTurnGroupIndex + 1) % data.groups.length;
      const isNewRound = nextIndex === 0;
      const finalStates = isNewRound ? startNewRound(newGroupStates) : newGroupStates;
      return {
        ...state,
        groupStates: finalStates,
        answerRevealed: false,
        openStationId: null,
        currentTurnGroupIndex: nextIndex,
        roundNumber: isNewRound ? state.roundNumber + 1 : state.roundNumber,
      };
    }
    case 'NEXT_TURN': {
      const nextIndex = (state.currentTurnGroupIndex + 1) % data.groups.length;
      const isNewRound = nextIndex === 0;
      const finalStates = isNewRound ? startNewRound(state.groupStates) : state.groupStates;
      return {
        ...state,
        groupStates: finalStates,
        openStationId: null,
        answerRevealed: false,
        currentTurnGroupIndex: nextIndex,
        roundNumber: isNewRound ? state.roundNumber + 1 : state.roundNumber,
      };
    }
    default:
      return state;
  }
}

export default function App() {
  const [state, dispatch] = useReducer(reducer, undefined, initState);
  const ui = data.uiCopy;

  const currentGroup = data.groups[state.currentTurnGroupIndex];

  const callbacks = {
    onOpenStation: (id: string) => dispatch({ type: 'OPEN_STATION', stationId: id }),
    onCloseStation: () => dispatch({ type: 'CLOSE_STATION' }),
    onRevealAnswer: () => dispatch({ type: 'REVEAL_ANSWER' }),
    onRecordOutcome: (groupId: string, outcome: AnswerOutcome) =>
      dispatch({ type: 'RECORD_OUTCOME', groupId, outcome }),
    onNextTurn: () => dispatch({ type: 'NEXT_TURN' }),
  };

  return (
    <div style={{ height: '100vh', display: 'flex', flexDirection: 'column', direction: 'rtl' }}>
      {/* Nav */}
      <nav style={NAV}>
        <span style={{ color: '#A0A8D8', fontSize: 11, padding: '4px 6px' }}>
          {ui.nav.screenLabel}
        </span>
        <NavBtn
          label={ui.nav.facilitator}
          active={state.view === 'facilitator'}
          onClick={() => dispatch({ type: 'SET_VIEW', view: 'facilitator' })}
        />
        {data.groups.map((g, i) => (
          <NavBtn
            key={g.id}
            label={`קב׳ ${i + 1} — ${g.divisionName}`}
            active={state.view === `mobile_${g.id}`}
            onClick={() => dispatch({ type: 'SET_VIEW', view: `mobile_${g.id}` })}
          />
        ))}
      </nav>

      {state.view === 'facilitator' ? (
        <FacilitatorView
          groupStates={state.groupStates}
          openStationId={state.openStationId}
          answerRevealed={state.answerRevealed}
          currentTurnGroupIndex={state.currentTurnGroupIndex}
          roundNumber={state.roundNumber}
          currentGroupId={currentGroup.id}
          {...callbacks}
        />
      ) : (
        <MobileView
          groupId={state.view.replace('mobile_', '')}
          currentTurnGroupId={currentGroup.id}
        />
      )}
    </div>
  );
}

function NavBtn({ label, active, onClick }: { label: string; active: boolean; onClick: () => void }) {
  return (
    <button onClick={onClick} style={{
      color: active ? '#000' : '#C8CEF0',
      fontSize: 12, padding: '4px 10px', borderRadius: 4,
      border: `1px solid ${active ? '#F5C800' : '#3A3E88'}`,
      background: active ? '#F5C800' : 'transparent',
      fontFamily: 'inherit', fontWeight: active ? 700 : 400,
      whiteSpace: 'nowrap',
    }}>
      {label}
    </button>
  );
}

const NAV: React.CSSProperties = {
  position: 'fixed', top: 0, right: 0, left: 0,
  background: '#1E1F56', padding: '6px 14px',
  display: 'flex', gap: 6, zIndex: 999, alignItems: 'center', flexWrap: 'wrap',
  borderBottom: '3px solid #F5C800',
};
