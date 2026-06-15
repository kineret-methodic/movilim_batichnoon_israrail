import React, { useState } from 'react';
import data from '../data/gameContent.json';
import { GROUP_COLORS, GROUP_BG } from '../engine';

interface Props {
  groupId: string;
  currentTurnGroupId: string;
}

const ui = data.uiCopy;

export default function MobileView({ groupId, currentTurnGroupId }: Props) {
  const [activeTab, setActiveTab] = useState<'tasks' | 'compensation'>('tasks');

  const group = data.groups.find(g => g.id === groupId);
  if (!group) return <div style={{ padding: 20 }}>קבוצה לא נמצאה</div>;

  const color = GROUP_COLORS[group.colorToken] ?? '#666';
  const bg = GROUP_BG[group.colorToken] ?? '#EEE';
  const tasks = (data.visibleTasks as any)[groupId] ?? [];
  const comp = (data.compensationModels as any)[groupId] ?? null;
  const isMyTurn = groupId === currentTurnGroupId;

  return (
    <div style={{
      paddingTop: 43, display: 'flex', flexDirection: 'column',
      height: '100vh', background: '#F0F2F7',
    }}>
      {/* Group header */}
      <div style={{
        background: '#1E1F56', padding: '10px 16px',
        borderBottom: `3px solid ${color}`,
        flexShrink: 0,
      }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
          <div style={{ width: 14, height: 14, borderRadius: '50%', background: color }} />
          <span style={{ color: 'white', fontWeight: 800, fontSize: 15 }}>{group.label}</span>
          <span style={{ color: '#8A92CC', fontSize: 12 }}>|</span>
          <span style={{ color: '#C8CEF0', fontSize: 12 }}>{group.divisionName}</span>
          {isMyTurn && (
            <span style={{
              marginRight: 'auto', fontSize: 11, background: color, color: 'white',
              borderRadius: 4, padding: '2px 7px', fontWeight: 600,
            }}>
              {ui.statusBar.activeTurn}
            </span>
          )}
        </div>
      </div>

      {/* Intro message */}
      <div style={{
        background: bg, borderBottom: `1px solid ${color}30`,
        padding: '10px 16px', fontSize: 12, color: '#5A5E80', flexShrink: 0,
        lineHeight: 1.5,
      }}>
        {ui.messages.mobileIntro}
      </div>

      {/* Tabs */}
      <div style={{
        display: 'flex', background: 'white', borderBottom: '2px solid #E0E4F0', flexShrink: 0,
      }}>
        {(['tasks', 'compensation'] as const).map(tab => {
          const label = tab === 'tasks' ? ui.tabs.reviewTasks : ui.tabs.compensationModel;
          const isActive = activeTab === tab;
          return (
            <button key={tab} onClick={() => setActiveTab(tab)} style={{
              flex: 1, padding: '11px 8px', fontWeight: 700, fontSize: 13,
              border: 'none', borderBottom: isActive ? `3px solid ${color}` : '3px solid transparent',
              background: isActive ? bg : 'white', color: isActive ? color : '#8A90C0',
              cursor: 'pointer',
            }}>
              {label}
            </button>
          );
        })}
      </div>

      {/* Content */}
      <div style={{ flex: 1, overflow: 'auto', padding: '12px 14px' }}>
        {activeTab === 'tasks' ? (
          <TasksTab tasks={tasks} group={group} color={color} bg={bg} />
        ) : (
          <CompensationTab comp={comp} color={color} bg={bg} />
        )}
      </div>
    </div>
  );
}

function TasksTab({ tasks, group, color, bg }: { tasks: any[]; group: any; color: string; bg: string }) {
  if (!tasks.length) {
    return <EmptyMsg msg={ui.messages.mobileTasksFallback} />;
  }
  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
      {tasks.map((task: any) => (
        <TaskCard key={task.taskId} task={task} color={color} bg={bg} />
      ))}
    </div>
  );
}

function TaskCard({ task, color, bg }: { task: any; color: string; bg: string }) {
  const [open, setOpen] = useState(true);
  const isTop10 = task.classification === 'top10';
  const classLabel = isTop10 ? ui.classificationLabels.top10 : ui.classificationLabels.routine;
  const tf = ui.taskFields as Record<string, string>;
  const ts = ui.taskSections as Record<string, string>;

  return (
    <div style={{
      background: 'white', borderRadius: 10, overflow: 'hidden',
      border: '1px solid #D0D4F0', boxShadow: '0 1px 4px rgba(30,31,86,.06)',
    }}>
      {/* Card header */}
      <div
        onClick={() => setOpen(o => !o)}
        style={{
          background: bg, padding: '10px 14px', cursor: 'pointer',
          display: 'flex', alignItems: 'center', gap: 8,
          borderBottom: open ? `1px solid ${color}30` : 'none',
        }}
      >
        <span style={{
          fontSize: 11, fontWeight: 700, padding: '2px 7px', borderRadius: 4,
          background: isTop10 ? color : '#8A90C0', color: 'white',
        }}>
          {classLabel}
        </span>
        <span style={{ fontWeight: 700, fontSize: 13, color: '#1A1C50', flex: 1 }}>
          {task.taskName}
        </span>
        <span style={{ fontSize: 14, color }}>{open ? '▲' : '▼'}</span>
      </div>

      {open && (
        <div style={{ padding: '12px 14px', fontSize: 12 }}>
          {/* Strategic */}
          <SectionHeader title={ts.strategic} />
          <FieldRow label={tf.companyGoal} value={task.companyGoal} />
          <FieldRow label={tf.goal} value={task.goal} />
          <FieldRow label={tf.outcomeMetric} value={task.outcomeMetric} />

          {/* Details */}
          <SectionHeader title={ts.details} />
          <FieldRow label={tf.taskDescription} value={task.taskDescription} />
          <FieldRow label={tf.annualAchievement} value={task.annualAchievement} />
          <FieldRow label={tf.outputMetricType} value={task.outputMetricType} />

          {/* Planning */}
          <SectionHeader title={ts.planning} />
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '4px 12px', marginTop: 4 }}>
            {(['Q1', 'Q2', 'Q3', 'Q4'] as const).map(q => (
              <div key={q} style={{ padding: '6px 8px', background: '#F5F6FF', borderRadius: 6, marginBottom: 2 }}>
                <div style={{ fontWeight: 700, fontSize: 11, color: '#1E1F56', marginBottom: 2 }}>{q}</div>
                <div style={{ color: '#3A3E60', lineHeight: 1.3 }}>{task.milestones[q]}</div>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

function CompensationTab({ comp, color, bg }: { comp: any; color: string; bg: string }) {
  if (!comp) return <EmptyMsg msg={ui.messages.mobileNoCompensation} />;
  const cf = ui.compensationFields as Record<string, string>;

  return (
    <div style={{
      background: 'white', borderRadius: 10, overflow: 'hidden',
      border: '1px solid #D0D4F0', boxShadow: '0 1px 4px rgba(30,31,86,.06)',
    }}>
      <div style={{ background: bg, padding: '10px 14px', borderBottom: `1px solid ${color}30` }}>
        <div style={{ fontWeight: 800, fontSize: 14, color: '#1A1C50' }}>{comp.modelName}</div>
      </div>
      <div style={{ padding: '12px 14px', fontSize: 12 }}>
        <FieldRow label={cf.rewardMetric} value={comp.rewardMetric} />
        <FieldRow label={cf.formula} value={comp.formula} />
        <FieldRow label={cf.weight} value={comp.weight} />
        <FieldRow label={cf.performanceThresholds} value={comp.performanceThresholds} />
        <FieldRow label={cf.dataSource} value={comp.dataSource} />
        {comp.notes && (
          <div style={{
            marginTop: 10, background: '#FFF8E0', border: '1px solid #E0D080',
            borderRadius: 6, padding: '8px 10px', fontSize: 11, color: '#5A5000',
          }}>
            ⚠️ {comp.notes}
          </div>
        )}
      </div>
    </div>
  );
}

function SectionHeader({ title }: { title: string }) {
  return (
    <div style={{
      fontSize: 10, fontWeight: 700, color: '#8A90C0', textTransform: 'uppercase',
      letterSpacing: .5, margin: '12px 0 4px', borderBottom: '1px solid #E8EAF0', paddingBottom: 3,
    }}>
      {title}
    </div>
  );
}

function FieldRow({ label, value }: { label: string; value: string }) {
  return (
    <div style={{ display: 'flex', gap: 6, marginBottom: 5, lineHeight: 1.4 }}>
      <span style={{ color: '#8A90C0', flexShrink: 0, minWidth: 100 }}>{label}:</span>
      <span style={{ color: '#1A1C50', fontWeight: 500 }}>{value}</span>
    </div>
  );
}

function EmptyMsg({ msg }: { msg: string }) {
  return (
    <div style={{ textAlign: 'center', color: '#8A90C0', padding: '30px 0', fontSize: 13 }}>{msg}</div>
  );
}
