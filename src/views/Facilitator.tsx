import React, { useState } from 'react';
import data from '../data/gameContent.json';
import type { GroupState, AnswerOutcome } from '../data/gameContent.types';
import { GROUP_COLORS, GROUP_BG } from '../engine';
import Timer from '../components/Timer';

interface Props {
  groupStates: Record<string, GroupState>;
  openStationId: string | null;
  answerRevealed: boolean;
  currentTurnGroupIndex: number;
  roundNumber: number;
  currentGroupId: string;
  onOpenStation: (id: string) => void;
  onCloseStation: () => void;
  onRevealAnswer: () => void;
  onRecordOutcome: (groupId: string, outcome: AnswerOutcome) => void;
  onNextTurn: () => void;
}

const ui = data.uiCopy;
const stations = data.stations as typeof data.stations;

// Group stations into 3 rows by y value (sorted by order)
const ROW_Y = [80, 48, 16];
const rows = ROW_Y.map(y =>
  (stations as any[])
    .filter(s => s.boardPosition.y === y)
    .sort((a: any, b: any) => a.order - b.order)
);

export default function FacilitatorView({
  groupStates, openStationId, answerRevealed,
  currentTurnGroupIndex, roundNumber, currentGroupId,
  onOpenStation, onCloseStation, onRevealAnswer, onRecordOutcome, onNextTurn,
}: Props) {
  const currentGroup = data.groups[currentTurnGroupIndex];
  const currentGs = groupStates[currentGroup.id];
  const currentStation = stations.find(s => s.order === currentGs.currentStationOrder);

  const turnText = (ui.turnIndicator as string)
    .replace('{group}', currentGroup.label)
    .replace('{division}', currentGroup.divisionName)
    .replace('{station}', currentStation?.name ?? '-');

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100vh', paddingTop: 43 }}>
      {/* Turn indicator bar */}
      <div style={{
        background: '#1E1F56', color: '#F5C800',
        padding: '8px 20px', fontSize: 14, fontWeight: 700,
        display: 'flex', alignItems: 'center', gap: 16, flexShrink: 0,
        borderBottom: '1px solid #2E3080',
      }}>
        <span>{turnText}</span>
        <span style={{ marginRight: 'auto', color: '#8A92CC', fontSize: 12, fontWeight: 400 }}>
          {ui.statusBar.roundComplete.includes('סבב') ? `סבב ${roundNumber}` : `סבב ${roundNumber}`}
        </span>
        <button onClick={onNextTurn} style={{
          background: '#2A2E72', border: '1px solid #3A3E88', color: '#C8CEF0',
          borderRadius: 6, padding: '4px 12px', fontSize: 12,
        }}>
          דלגו על תור ←
        </button>
      </div>

      {/* Board + Status */}
      <div style={{ flex: 1, overflow: 'auto', padding: '10px 14px', display: 'flex', flexDirection: 'column', gap: 8 }}>
        <GameMap
          groupStates={groupStates}
          openStationId={openStationId}
          currentGroupId={currentGroupId}
          onOpenStation={onOpenStation}
        />
        <GroupStatusBar groupStates={groupStates} currentGroupId={currentGroupId} />
      </div>

      {/* Station Panel overlay */}
      {openStationId && (
        <StationPanel
          stationId={openStationId}
          groupStates={groupStates}
          answerRevealed={answerRevealed}
          currentGroupId={currentGroupId}
          onClose={onCloseStation}
          onRevealAnswer={onRevealAnswer}
          onRecordOutcome={onRecordOutcome}
        />
      )}
    </div>
  );
}

// ─── Game Map ────────────────────────────────────────────────────────────────

function GameMap({ groupStates, openStationId, currentGroupId, onOpenStation }: {
  groupStates: Record<string, GroupState>;
  openStationId: string | null;
  currentGroupId: string;
  onOpenStation: (id: string) => void;
}) {
  const sectionLabels = ui.boardSections as Record<string, string>;

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 0 }}>
      {/* Section label — compensation */}
      <SectionLabel title={sectionLabels.c} months={sectionLabels.cMonths} />

      {/* Row 3 (y=16) — RTL via flex row */}
      <BoardRow stations={rows[2]} ltr={false} groupStates={groupStates} openStationId={openStationId} currentGroupId={currentGroupId} onOpenStation={onOpenStation} />

      {/* Connector right→left (snake top-right) */}
      <Connector side="right" />

      {/* Section label — middle */}
      <SectionLabel title={sectionLabels.b} months={sectionLabels.bMonths} />

      {/* Row 2 (y=48) — LTR via flex row-reverse */}
      <BoardRow stations={rows[1]} ltr={true} groupStates={groupStates} openStationId={openStationId} currentGroupId={currentGroupId} onOpenStation={onOpenStation} />

      {/* Connector left→right (snake bottom-left) */}
      <Connector side="left" />

      {/* Section label — workplan */}
      <SectionLabel title={sectionLabels.a} months={sectionLabels.aMonths} />

      {/* Row 1 (y=80) — RTL */}
      <BoardRow stations={rows[0]} ltr={false} groupStates={groupStates} openStationId={openStationId} currentGroupId={currentGroupId} onOpenStation={onOpenStation} />
    </div>
  );
}

function SectionLabel({ title, months }: { title: string; months: string }) {
  return (
    <div style={{ display: 'flex', justifyContent: 'space-between', padding: '4px 4px 2px', marginTop: 6 }}>
      <span style={{ fontSize: 11, fontWeight: 600, color: '#1E1F56' }}>{title}</span>
      <span style={{ fontSize: 10, color: '#8A90C0' }}>{months}</span>
    </div>
  );
}

function Connector({ side }: { side: 'left' | 'right' }) {
  const isRight = side === 'right';
  return (
    <div style={{
      display: 'flex', flexDirection: 'column',
      alignItems: isRight ? 'flex-start' : 'flex-end',
      paddingInlineEnd: isRight ? 0 : 20,
      paddingInlineStart: isRight ? 20 : 0,
      height: 22,
    }}>
      <div style={{ width: 1.5, background: '#B0B4D8', flex: 1 }} />
      <span style={{ fontSize: 12, color: '#B0B4D8', lineHeight: 1 }}>▲</span>
    </div>
  );
}

function BoardRow({ stations, ltr, groupStates, openStationId, currentGroupId, onOpenStation }: {
  stations: any[];
  ltr: boolean;
  groupStates: Record<string, GroupState>;
  openStationId: string | null;
  currentGroupId: string;
  onOpenStation: (id: string) => void;
}) {
  return (
    <div style={{
      display: 'flex',
      flexDirection: ltr ? 'row-reverse' : 'row',
      alignItems: 'center',
      gap: 0,
      background: '#FFFFFF',
      borderRadius: 12,
      padding: '8px 10px',
      border: '1px solid #D0D4F0',
      boxShadow: '0 1px 4px rgba(30,31,86,.06)',
      overflowX: 'auto',
    }}>
      {stations.map((station, i) => (
        <React.Fragment key={station.id}>
          {i > 0 && (
            <span style={{ fontSize: 12, color: '#B0B4D8', flexShrink: 0, padding: '0 3px' }}>
              {ltr ? '←' : '→'}
            </span>
          )}
          <StationNode
            station={station}
            groupStates={groupStates}
            isOpen={openStationId === station.id}
            currentGroupId={currentGroupId}
            onOpen={() => onOpenStation(station.id)}
          />
        </React.Fragment>
      ))}
    </div>
  );
}

function StationNode({ station, groupStates, isOpen, currentGroupId, onOpen }: {
  station: any;
  groupStates: Record<string, GroupState>;
  isOpen: boolean;
  currentGroupId: string;
  onOpen: () => void;
}) {
  const groupsHere = data.groups.filter(g => groupStates[g.id]?.currentStationOrder === station.order);
  const isCurrentGroupHere = groupStates[currentGroupId]?.currentStationOrder === station.order;

  const isUE = station.type === 'unexpected_event';
  const isTr = station.type === 'transition';
  const isComp = station.segment === 'compensation';
  const isPro = station.type === 'professional';
  const isRefl = station.type === 'reflection';

  let bg = '#FFFFFF';
  let border = '1.5px solid #1E1F56';
  if (isUE)   { bg = '#FFF3E0'; border = '1.5px dashed #E07820'; }
  else if (isRefl) { bg = '#F5F0FF'; border = '1.5px dashed #8A70C0'; }
  else if (isTr && isComp) { bg = '#E8F5EE'; border = '2px solid #1A6A3A'; }
  else if (isTr) { bg = '#EEF0FF'; border = '1.5px solid #5A5EAA'; }
  else if (isComp) { bg = '#F0FFF8'; border = '1.5px solid #1A6A3A'; }

  if (isOpen) border = '2.5px solid #F5C800';
  if (isCurrentGroupHere && !isOpen) border = '2.5px solid #2A7AE4';

  return (
    <div
      onClick={onOpen}
      data-station-id={station.id}
      style={{
        borderRadius: isUE ? '50%' : 8,
        padding: isUE ? '8px 6px' : '7px 8px',
        minWidth: isUE ? 64 : 82,
        maxWidth: isUE ? 72 : 108,
        textAlign: 'center',
        flexShrink: 0,
        cursor: 'pointer',
        background: bg,
        border,
        boxShadow: isOpen ? '0 0 0 3px #F5C80066' : '0 1px 3px rgba(30,31,86,.1)',
        transition: 'box-shadow .15s, transform .1s',
        position: 'relative',
      }}
      onMouseEnter={e => (e.currentTarget.style.transform = 'scale(1.05)')}
      onMouseLeave={e => (e.currentTarget.style.transform = 'scale(1)')}
    >
      {station.month && (
        <div style={{ fontSize: 8, color: '#8A90C0', marginBottom: 2 }}>{station.month}</div>
      )}
      <div style={{ fontSize: isUE ? 9 : 10, fontWeight: 700, color: isUE ? '#7A3A00' : '#1A1C50', lineHeight: 1.3 }}>
        {station.name}
      </div>

      {/* Group tokens */}
      {groupsHere.length > 0 && (
        <div style={{ display: 'flex', justifyContent: 'center', gap: 3, marginTop: 4, flexWrap: 'wrap' }}>
          {groupsHere.map(g => (
            <div key={g.id} style={{
              width: 10, height: 10, borderRadius: '50%',
              background: GROUP_COLORS[g.colorToken] ?? '#999',
              border: '1.5px solid white',
              boxShadow: '0 1px 2px rgba(0,0,0,.3)',
            }} title={g.label} />
          ))}
        </div>
      )}
    </div>
  );
}

// ─── Group Status Bar ─────────────────────────────────────────────────────────

function GroupStatusBar({ groupStates, currentGroupId }: {
  groupStates: Record<string, GroupState>;
  currentGroupId: string;
}) {
  return (
    <div style={{ display: 'flex', gap: 8 }}>
      {data.groups.map(g => {
        const gs = groupStates[g.id];
        const station = stations.find(s => s.order === gs.currentStationOrder);
        const isCurrent = g.id === currentGroupId;
        const color = GROUP_COLORS[g.colorToken] ?? '#999';
        const bg = GROUP_BG[g.colorToken] ?? '#EEE';

        let statusText = ui.statusBar.notYet;
        if (isCurrent) statusText = ui.statusBar.activeTurn;
        else if (gs.waitRoundsRemaining > 0) statusText = ui.statusBar.waiting;
        else if (gs.pendingAutoAdvance) statusText = ui.statusBar.autoAdvance;

        return (
          <div key={g.id} style={{
            flex: 1, background: isCurrent ? bg : '#FFFFFF',
            border: `1.5px solid ${isCurrent ? color : '#D0D4F0'}`,
            borderRadius: 10, padding: '7px 10px',
          }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 6, marginBottom: 4 }}>
              <div style={{ width: 10, height: 10, borderRadius: '50%', background: color }} />
              <span style={{ fontSize: 12, fontWeight: 700, color }}>{g.label}</span>
              {isCurrent && (
                <span style={{ fontSize: 10, background: color, color: 'white', borderRadius: 4, padding: '1px 5px', marginRight: 'auto' }}>
                  {ui.statusBar.activeTurn}
                </span>
              )}
            </div>
            <div style={{ fontSize: 11, color: '#5A5E80', marginBottom: 2 }}>
              📍 {station?.name ?? '-'}
            </div>
            <div style={{ display: 'flex', gap: 10, fontSize: 11 }}>
              <span>🏆 {gs.points} {ui.points.label}</span>
              <span>⚠️ {gs.errors} טעויות</span>
            </div>
            {!isCurrent && (
              <div style={{ fontSize: 10, color: '#8A90C0', marginTop: 2 }}>{statusText}</div>
            )}
          </div>
        );
      })}
    </div>
  );
}

// ─── Station Panel (overlay) ──────────────────────────────────────────────────

function StationPanel({ stationId, groupStates, answerRevealed, currentGroupId, onClose, onRevealAnswer, onRecordOutcome }: {
  stationId: string;
  groupStates: Record<string, GroupState>;
  answerRevealed: boolean;
  currentGroupId: string;
  onClose: () => void;
  onRevealAnswer: () => void;
  onRecordOutcome: (groupId: string, outcome: AnswerOutcome) => void;
}) {
  const station = stations.find(s => s.id === stationId);
  const [activeTab, setActiveTab] = useState(currentGroupId);
  const [reminderOpen, setReminderOpen] = useState(false);

  if (!station) return null;

  const question = station.questionId
    ? (data.stationQuestions as any[]).find(q => q.id === station.questionId)
    : null;

  const transMsg = station.transitionMessageId
    ? (data.transitionMessages as any)[station.transitionMessageId]
    : null;

  const isUE = station.type === 'unexpected_event';
  const isTr = station.type === 'transition';

  return (
    <div style={{
      position: 'fixed', inset: 0, zIndex: 500,
      background: 'rgba(10,12,40,.75)',
      display: 'flex', alignItems: 'flex-end',
    }} onClick={e => { if (e.target === e.currentTarget) onClose(); }}>
      <div style={{
        width: '100%', background: '#FFFFFF',
        borderRadius: '16px 16px 0 0',
        boxShadow: '0 -4px 24px rgba(30,31,86,.2)',
        maxHeight: '88vh', display: 'flex', flexDirection: 'column',
        overflow: 'hidden',
      }}>
        {/* Panel header */}
        <div style={{
          background: '#1E1F56', color: 'white',
          padding: '10px 18px', display: 'flex', alignItems: 'center', gap: 12, flexShrink: 0,
        }}>
          <button onClick={onClose} style={{
            background: 'none', border: '1px solid #3A3E88', color: '#C8CEF0',
            borderRadius: 6, padding: '4px 10px', fontSize: 12,
          }}>
            {ui.buttons.backToMap}
          </button>
          <span style={{ fontWeight: 800, fontSize: 17 }}>{station.name}</span>
          {station.month && (
            <span style={{ fontSize: 12, color: '#8A92CC', marginRight: 'auto' }}>{station.month}</span>
          )}
        </div>

        <div style={{ flex: 1, overflow: 'auto', padding: 18 }}>
          {/* ── Transition station ── */}
          {isTr && transMsg && (
            <div style={{ textAlign: 'center', padding: '24px 0' }}>
              <div style={{ fontSize: 22, fontWeight: 800, color: '#1E1F56', marginBottom: 12 }}>{transMsg.title}</div>
              <div style={{ fontSize: 15, color: '#3A3E60', maxWidth: 600, margin: '0 auto 20px' }}>{transMsg.body}</div>
              <button onClick={onClose} style={outcomeBtn('#1E1F56', 'white')}>
                {ui.groupCard.approveTransition}
              </button>
            </div>
          )}

          {/* ── Unexpected event station ── */}
          {isUE && (
            <div style={{ textAlign: 'center', padding: '20px 0' }}>
              <div style={{ fontSize: 40, marginBottom: 10 }}>🎰</div>
              <div style={{ fontSize: 18, fontWeight: 700, color: '#7A3A00', marginBottom: 8 }}>
                {ui.stationTypes.unexpectedEventLabel}
              </div>
              <div style={{ fontSize: 13, color: '#5A5E80', marginBottom: 16 }}>
                {ui.buttons.spinWheel} — בלת״מ זה בשלב האבטיפוס
              </div>
              <button onClick={onClose} style={outcomeBtn('#E07820', 'white')}>
                {ui.buttons.backToMap}
              </button>
            </div>
          )}

          {/* ── Professional station ── */}
          {question && (
            <>
              {/* Question */}
              <div style={{
                background: '#F5F6FF', border: '1.5px solid #D0D4F0',
                borderRadius: 10, padding: '14px 16px', marginBottom: 14,
              }}>
                <div style={{ fontSize: 15, fontWeight: 700, color: '#1A1C50', marginBottom: 8, lineHeight: 1.5 }}>
                  {question.questionText}
                </div>
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: 8 }}>
                  <Timer
                    initialSeconds={question.timerSeconds}
                    stopLabel={ui.timer.stop}
                    resumeLabel={ui.timer.resume}
                  />
                  <button
                    onClick={() => setReminderOpen(r => !r)}
                    style={{ fontSize: 12, color: '#3A5EAA', background: 'none', border: '1px solid #C0C4E8', borderRadius: 6, padding: '4px 10px' }}
                  >
                    {ui.buttons.needReminder}
                  </button>
                </div>
              </div>

              {/* Reminder drawer */}
              {reminderOpen && (
                <div style={{
                  background: '#FFFFF0', border: '1px solid #E0D080', borderRadius: 8,
                  padding: '10px 14px', marginBottom: 12, fontSize: 12, color: '#5A5000',
                }}>
                  <div style={{ fontWeight: 700, marginBottom: 6 }}>{ui.reminderDrawer.title}</div>
                  {(data.reminderDefinitions as any[]).map((r: any) => (
                    <div key={r.term} style={{ marginBottom: 4 }}>
                      <strong>{r.term}:</strong> {r.definition}
                    </div>
                  ))}
                </div>
              )}

              {/* Principle (for facilitator, subtle) */}
              {answerRevealed && (
                <div style={{
                  background: '#EFF8EE', border: '1px solid #B0D8B0',
                  borderRadius: 8, padding: '8px 12px', marginBottom: 12,
                  fontSize: 12, color: '#1A5A1A',
                }}>
                  <strong>עיקרון מקצועי: </strong>{question.professionalPrinciple}
                </div>
              )}

              {/* Group tabs */}
              <div style={{ display: 'flex', gap: 4, marginBottom: 0, borderBottom: '2px solid #E0E4F0' }}>
                {data.groups.map(g => {
                  const color = GROUP_COLORS[g.colorToken];
                  const isActive = activeTab === g.id;
                  const gs = groupStates[g.id];
                  const isAtStation = gs.currentStationOrder === station.order;
                  return (
                    <button key={g.id} onClick={() => setActiveTab(g.id)} style={{
                      padding: '8px 14px', fontWeight: 700, fontSize: 13,
                      border: 'none', borderBottom: isActive ? `3px solid ${color}` : '3px solid transparent',
                      background: isActive ? GROUP_BG[g.colorToken] : 'none',
                      color: isActive ? color : '#8A90C0',
                      cursor: 'pointer', borderRadius: '6px 6px 0 0',
                      display: 'flex', alignItems: 'center', gap: 5,
                    }}>
                      <span style={{
                        width: 8, height: 8, borderRadius: '50%', background: color, display: 'inline-block',
                        opacity: isAtStation ? 1 : 0.35,
                      }} />
                      {g.label}
                      {isAtStation && (
                        <span style={{ fontSize: 9, background: color, color: 'white', borderRadius: 3, padding: '1px 4px' }}>
                          כאן
                        </span>
                      )}
                    </button>
                  );
                })}
              </div>

              {/* Active tab content */}
              <GroupTabContent
                groupId={activeTab}
                question={question}
                station={station}
                groupStates={groupStates}
                answerRevealed={answerRevealed}
                currentGroupId={currentGroupId}
                onRevealAnswer={onRevealAnswer}
                onRecordOutcome={onRecordOutcome}
              />
            </>
          )}
        </div>
      </div>
    </div>
  );
}

function GroupTabContent({ groupId, question, station, groupStates, answerRevealed, currentGroupId, onRevealAnswer, onRecordOutcome }: {
  groupId: string;
  question: any;
  station: any;
  groupStates: Record<string, GroupState>;
  answerRevealed: boolean;
  currentGroupId: string;
  onRevealAnswer: () => void;
  onRecordOutcome: (groupId: string, outcome: AnswerOutcome) => void;
}) {
  const gs = groupStates[groupId];
  const isAtStation = gs?.currentStationOrder === station.order;
  const answer = question.answersByGroup[groupId];
  const isCurrent = groupId === currentGroupId;
  const g = data.groups.find(gr => gr.id === groupId)!;
  const color = GROUP_COLORS[g.colorToken];

  return (
    <div style={{ padding: '14px 0' }}>
      {!isAtStation && (
        <div style={{ fontSize: 13, color: '#8A90C0', padding: '8px 0', fontStyle: 'italic' }}>
          {ui.groupCard.stationLocked}
        </div>
      )}

      {/* Reveal / Answer */}
      {isAtStation && !answerRevealed && (
        <button onClick={onRevealAnswer} style={{
          ...outcomeBtn('#1E1F56', 'white'), marginBottom: 14,
        }}>
          {ui.buttons.revealAnswer}
        </button>
      )}

      {answerRevealed && answer && (
        <div style={{
          background: '#F8F9FF', border: `1.5px solid ${color}30`,
          borderRight: `4px solid ${color}`,
          borderRadius: 8, padding: '12px 14px', marginBottom: 14,
        }}>
          <div style={{ fontSize: 12, color: '#8A90C0', marginBottom: 4 }}>{ui.answerCard.taskLabel}</div>
          <div style={{ fontWeight: 700, fontSize: 14, color: '#1A1C50', marginBottom: 10 }}>{answer.taskName}</div>

          <div style={{ fontSize: 12, color: '#8A90C0', marginBottom: 2 }}>{ui.answerCard.whatIsWrongLabel}</div>
          <div style={{ fontSize: 13, color: '#CC3300', marginBottom: 10 }}>{answer.whatIsWrong}</div>

          <div style={{ fontSize: 12, color: '#8A90C0', marginBottom: 2 }}>{ui.answerCard.correctFixLabel}</div>
          <div style={{ fontSize: 13, color: '#1A5A1A', fontWeight: 600 }}>{answer.correctFix}</div>

          {answer.facilitatorNotes && (
            <>
              <div style={{ fontSize: 12, color: '#8A90C0', marginTop: 8, marginBottom: 2 }}>{ui.answerCard.facilitatorNoteLabel}</div>
              <div style={{ fontSize: 12, color: '#5A5E80', fontStyle: 'italic' }}>{answer.facilitatorNotes}</div>
            </>
          )}
        </div>
      )}

      {answerRevealed && !answer && (
        <div style={{ fontSize: 13, color: '#8A90C0', fontStyle: 'italic', marginBottom: 14 }}>
          {ui.messages.noGroupAnswer}
        </div>
      )}

      {/* Outcome buttons */}
      {isAtStation && (
        <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
          <button
            onClick={() => { if (answerRevealed) onRecordOutcome(groupId, 'correct'); }}
            style={outcomeBtn('#1A7A3A', 'white', !answerRevealed)}
          >
            ✓ {ui.buttons.correct}
          </button>
          <button
            onClick={() => { if (answerRevealed) onRecordOutcome(groupId, 'partial'); }}
            style={outcomeBtn('#C05A10', 'white', !answerRevealed)}
          >
            ◑ {ui.buttons.partial}
          </button>
          <button
            onClick={() => { if (answerRevealed) onRecordOutcome(groupId, 'wrong'); }}
            style={outcomeBtn('#CC2200', 'white', !answerRevealed)}
          >
            ✗ {ui.buttons.wrong}
          </button>
        </div>
      )}

      {/* Outcome notes */}
      {answerRevealed && isAtStation && (
        <div style={{ marginTop: 10, fontSize: 11, color: '#8A90C0', display: 'flex', flexDirection: 'column', gap: 2 }}>
          <span>✓ נכון → {ui.messages.correctOutcome}</span>
          <span>◑ חלקי → {ui.messages.partialOutcome}</span>
          <span>✗ שגוי → {ui.messages.wrongOutcome}</span>
        </div>
      )}
    </div>
  );
}

function outcomeBtn(bg: string, color: string, disabled = false): React.CSSProperties {
  return {
    padding: '9px 18px', borderRadius: 8, border: 'none',
    background: bg, color, fontSize: 14, fontWeight: 700,
    cursor: disabled ? 'not-allowed' : 'pointer',
    opacity: disabled ? 0.4 : 1,
    transition: 'opacity .15s',
  };
}
