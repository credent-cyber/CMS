/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable  @typescript-eslint/no-explicit-any*/
/* eslint-disable no-void*/
import * as React from 'react';
import { useRef, useState, useEffect, useMemo, useCallback } from 'react';
import { DataGrid, GridColDef, GridRenderCellParams } from '@mui/x-data-grid';
import {
  Chip, Tooltip, Switch,
  Select, MenuItem, SelectChangeEvent,
  Drawer, IconButton, Button,
   Dialog,
  DialogTitle,
  DialogContent,
  DialogActions,
  RadioGroup, FormControlLabel, Radio,
} from '@mui/material';
import AddCircleOutlineIcon  from '@mui/icons-material/AddCircleOutline';
import QuizOutlinedIcon      from '@mui/icons-material/QuizOutlined';
import KeyboardArrowDownIcon from '@mui/icons-material/KeyboardArrowDown';
import SyncIcon              from '@mui/icons-material/Sync';
import ErrorOutlineIcon      from '@mui/icons-material/ErrorOutline';
import WarningAmberIcon      from '@mui/icons-material/WarningAmber';
import SearchIcon            from '@mui/icons-material/Search';
import CloseIcon             from '@mui/icons-material/Close';
import CheckCircleOutlineIcon from '@mui/icons-material/CheckCircleOutline';
import EditOutlinedIcon       from '@mui/icons-material/EditOutlined';
import DeleteOutlineIcon       from '@mui/icons-material/DeleteOutline';
import ArrowUpwardIcon         from '@mui/icons-material/ArrowUpward';
import ArrowDownwardIcon       from '@mui/icons-material/ArrowDownward';
import { sp }                from '@pnp/sp';
import { useSnackbar }       from './Snackbar';
import { BuddyLoader }       from './Buddyloader';
import LIST_CONFIG           from '../../../config/spListConfig';
import styles                from './BuddyQuestionnaire.module.scss';
import type { IBuddyAppProps } from './IBuddyAppProps';

interface Question {
  id: string;
  spId: number;
  meetingNo: number;
  meeting: string;
  text: string;
  status: SPStatus;
}

type SPStatus = 'Active' | 'InActive';
type RowSaveState = 'idle' | 'saving';
type StatusFilter = 'All' | 'Active' | 'InActive';


const LIST_NAME = LIST_CONFIG.LISTS.BUDDY_QUESTIONS ?? 'Buddy Interaction Questions Master';

const MEETING_OPTIONS: { value: number; label: string }[] = [
  { value: 1, label: '1st Meeting' },
  { value: 2, label: '2nd Meeting' },
  { value: 3, label: '3rd Meeting' },
  { value: 4, label: '4th Meeting' },
  { value: 5, label: '5th Meeting' },
  { value: 6, label: '6th Meeting' },
];

const labelToNo = (label: string): number =>
  MEETING_OPTIONS.find(m => m.label === label)?.value ?? 1;


const DRAWER_OVER_TOPBAR_SX = {
  zIndex: 99999,
  '& .MuiBackdrop-root': {
    zIndex: 99998,
  },
  '& .MuiDrawer-paper': {
    top: '0 !important',
    height: '100vh !important',
    maxHeight: '100vh !important',
    zIndex: 100000,
  },
};


export const BuddyQuestionnaire: React.FC<IBuddyAppProps> = (props) => {
  const isMountedRef     = useRef(true);
  const { showSnackbar } = useSnackbar();
  const textareaRef      = useRef<HTMLTextAreaElement>(null);

  const [questions,        setQuestions]        = useState<Question[]>([]);
  const [isLoading,        setIsLoading]        = useState(false);
  const [loadError,        setLoadError]        = useState('');
  const [selectedMeeting,  setSelectedMeeting]  = useState<number>(1);
  const [selectedStatus,   setSelectedStatus]   = useState<StatusFilter>('All');
  const [questionSortDirection, setQuestionSortDirection] = useState<'asc' | 'desc'>('asc');
  const [isAdding,         setIsAdding]         = useState<boolean>(false);
  const [searchText,       setSearchText]       = useState<string>('');
  const [rowSaveStates,    setRowSaveStates]    = useState<Record<string, RowSaveState>>({});
  const [loaderVisible,    setLoaderVisible]    = useState<boolean>(false);
  const [loaderMessage,    setLoaderMessage]    = useState<string>('Processing your request…');

  const [drawerOpen,          setDrawerOpen]          = useState(false);
  const [drawerMeeting,       setDrawerMeeting]       = useState<number>(1);
  const [drawerText,          setDrawerText]          = useState<string>('');
  const [drawerDupWarning,    setDrawerDupWarning]    = useState<string>('');
  const [drawerCharCount,     setDrawerCharCount]     = useState<number>(0);
  const [drawerSuccess,       setDrawerSuccess]       = useState<boolean>(false);
  const [editingQuestion,     setEditingQuestion]     = useState<Question | null>(null);
  const [deleteTarget,        setDeleteTarget]        = useState<Question | null>(null);

  useEffect(() => { return () => { isMountedRef.current = false; }; }, []);

  useEffect(() => {
    if (drawerOpen) {
      setDrawerMeeting(editingQuestion?.meetingNo ?? selectedMeeting);
      setDrawerText(editingQuestion?.text ?? '');
      setDrawerDupWarning('');
      setDrawerCharCount(editingQuestion?.text.length ?? 0);
      setDrawerSuccess(false);
      setTimeout(() => { textareaRef.current?.focus(); }, 320);
    }
  }, [drawerOpen, editingQuestion, selectedMeeting]);


  const loadQuestions = useCallback(async () => {
    if (!props.context) return;
    setIsLoading(true);
    setLoadError('');
    setLoaderVisible(true);
    setLoaderMessage('Loading questions…');
    try {
      const items = await sp.web.lists
        .getByTitle(LIST_NAME)
        .items
        .select('ID,Title,Meeting,Questions,Status')
        .orderBy('Meeting')
        .orderBy('ID')
        
        .top(5000)
        .get();

      if (!isMountedRef.current) return;

      setQuestions(
        (items || []).map((i: any) => ({
          id:        i.ID.toString(),
          spId:      i.ID,
          meeting:   i.Meeting   || '1st Meeting',
          meetingNo: labelToNo(i.Meeting || '1st Meeting'),
          text:      i.Questions || i.Title || '',
          status:    (i.Status === 'Active' ? 'Active' : 'InActive') as SPStatus,
        })),
      );
    } catch (e) {
      console.error('loadQuestions:', e);
      if (isMountedRef.current) setLoadError('Failed to load questions. Please refresh.');
    } finally {
      if (isMountedRef.current) {
        setIsLoading(false);
        setLoaderVisible(false);
      }
    }
  }, [props.context]);

  useEffect(() => { void loadQuestions(); }, [loadQuestions]);


  const openAddQuestionDrawer = () => {
    setEditingQuestion(null);
    setDrawerOpen(true);
  };

  const openEditQuestionDrawer = (question: Question) => {
    setEditingQuestion(question);
    setDrawerOpen(true);
  };

  const handleSaveQuestion = async () => {
    const trimmed = drawerText.trim();
    if (!trimmed || isAdding) return;

    const meetingLabel = MEETING_OPTIONS.find(m => m.value === drawerMeeting)?.label ?? '1st Meeting';

    const isDuplicate = questions.some(
      q => q.id !== editingQuestion?.id &&
           q.meetingNo === drawerMeeting &&
           q.text.trim().toLowerCase() === trimmed.toLowerCase(),
    );

    if (isDuplicate) {
      setDrawerDupWarning(
        `This question already exists for ${meetingLabel}. Please enter a unique question.`,
      );
      return;
    }

    setDrawerDupWarning('');
    setIsAdding(true);
    setLoaderVisible(!editingQuestion);
    if (!editingQuestion) setLoaderMessage('Adding new question…');

    try {
      if (editingQuestion) {
        await sp.web.lists
          .getByTitle(LIST_NAME)
          .items.getById(editingQuestion.spId)
          .update({
            Title:     trimmed.substring(0, 255),
            Meeting:   meetingLabel,
            Questions: trimmed,
          });

        if (!isMountedRef.current) return;

        setQuestions(prev => prev.map(item =>
          item.id === editingQuestion.id
            ? {
                ...item,
                meeting: meetingLabel,
                meetingNo: drawerMeeting,
                text: trimmed,
              }
            : item,
        ));
        setSelectedMeeting(drawerMeeting);
        setSearchText('');
        setDrawerSuccess(true);
        showSnackbar('The buddy meeting question has been updated successfully.', 'success', 3000);

        setTimeout(() => {
          if (isMountedRef.current) {
            setDrawerOpen(false);
            setDrawerText('');
            setDrawerSuccess(false);
            setEditingQuestion(null);
          }
        }, 1100);
        return;
      }

      const result = await sp.web.lists
        .getByTitle(LIST_NAME)
        .items.add({
          Title:     trimmed.substring(0, 255),
          Meeting:   meetingLabel,
          Questions: trimmed,
          Status:    'Active',
        });

      if (!isMountedRef.current) return;

      const newId = result.data.ID;
      setQuestions(prev => [...prev, {
        id:        newId.toString(),
        spId:      newId,
        meeting:   meetingLabel,
        meetingNo: drawerMeeting,
        text:      trimmed,
        status:    'Active',
      }]);

      setSelectedMeeting(drawerMeeting);
      setSearchText('');

      setDrawerSuccess(true);
      showSnackbar(`Question added to ${meetingLabel} successfully!`, 'success', 3000);

      setTimeout(() => {
        if (isMountedRef.current) {
          setDrawerOpen(false);
          setDrawerText('');
          setDrawerSuccess(false);
        }
      }, 1400);

    } catch (e) {
      console.error(editingQuestion ? 'updateQuestion:' : 'addQuestion:', e);
      showSnackbar(
        editingQuestion
          ? 'Failed to update question. Please try again.'
          : 'Failed to add question. Please try again.',
        'error',
        4000,
      );
    } finally {
      if (isMountedRef.current) {
        setIsAdding(false);
        setLoaderVisible(false);
      }
    }
  };

  const handleDrawerClose = () => {
    if (isAdding) return;
    setDrawerOpen(false);
    setEditingQuestion(null);
    setDrawerText('');
    setDrawerDupWarning('');
    setDrawerSuccess(false);
  };


  const handleToggleStatus = async (q: Question) => {
    const nextStatus: SPStatus = q.status === 'Active' ? 'InActive' : 'Active';

    setQuestions(prev => prev.map(item =>
      item.id === q.id ? { ...item, status: nextStatus } : item,
    ));
    setRowSaveStates(prev => ({ ...prev, [q.id]: 'saving' }));
    setLoaderVisible(true);
    setLoaderMessage('Updating question status…');

    try {
      await sp.web.lists
        .getByTitle(LIST_NAME)
        .items.getById(q.spId)
        .update({ Status: nextStatus });

      if (!isMountedRef.current) return;

      showSnackbar(
        nextStatus === 'Active' ? 'Question marked as Active.' : 'Question marked as InActive.',
        nextStatus === 'Active' ? 'success' : 'warning',
        3000,
      );
    } catch (e) {
      console.error('toggleStatus:', e);
      if (!isMountedRef.current) return;
      setQuestions(prev => prev.map(item =>
        item.id === q.id ? { ...item, status: q.status } : item,
      ));
      showSnackbar('Failed to update status. Please try again.', 'error', 4000);
    } finally {
      if (isMountedRef.current) {
        setRowSaveStates(prev => ({ ...prev, [q.id]: 'idle' }));
        setLoaderVisible(false);
      }
    }
  };


  const handleDeleteQuestion = (q: Question) => {
    setDeleteTarget(q);
  };

  const handleConfirmDeleteQuestion = async () => {
    if (!deleteTarget) return;
    const q = deleteTarget;
    setDeleteTarget(null);

    setRowSaveStates(prev => ({ ...prev, [q.id]: 'saving' }));
    setLoaderVisible(true);
    setLoaderMessage('Removing question…');

    try {
      await sp.web.lists
        .getByTitle(LIST_NAME)
        .items.getById(q.spId)
        .delete();

      if (!isMountedRef.current) return;

      setQuestions(prev => prev.filter(item => item.id !== q.id));
      showSnackbar('Question removed successfully.', 'success', 3000);
    } catch (e) {
      console.error('deleteQuestion:', e);
      showSnackbar("We couldn't remove the question. Please try again.", 'error', 4000);
    } finally {
      if (isMountedRef.current) {
        setRowSaveStates(prev => ({ ...prev, [q.id]: 'idle' }));
        setLoaderVisible(false);
      }
    }
  };

  const renderDeleteConfirmation = () => {
  if (!deleteTarget) return null;

  return (
    <Dialog
      open={Boolean(deleteTarget)}
      onClose={() => setDeleteTarget(null)}
      maxWidth="sm"
      fullWidth
      PaperProps={{
        sx: {
          borderRadius: '18px',
          overflow: 'hidden',
          boxShadow: '0 20px 55px rgba(0, 0, 0, 0.25)',
        },
      }}
    >
      <DialogTitle
        sx={{
          backgroundColor: '#004632',
          color: '#ffffff',
          fontSize: '1.25rem',
          fontWeight: 700,
          px: 4,
          py: 2.5,
        }}
      >
        Confirm Question Removal
      </DialogTitle>

      <DialogContent
        sx={{
          px: 4,
          pt: '22px !important',
          pb: 2,
          backgroundColor: '#ffffff',
        }}
      >
        <div
          style={{
            fontSize: '0.95rem',
            lineHeight: 1.7,
            color: '#374151',
          }}
        >
          Are you sure you want to remove this question?
        </div>

        {/* <div
          style={{
            marginTop: '10px',
            padding: '12px 14px',
            borderRadius: '10px',
            backgroundColor: '#f8faf9',
            border: '1px solid #d9e7e1',
            color: '#1f2937',
            fontSize: '0.875rem',
            fontWeight: 600,
          }}
        >
          “{deleteTarget.text}”
        </div> */}

        {/* <div
          style={{
            marginTop: '14px',
            fontSize: '0.82rem',
            lineHeight: 1.6,
            color: '#6b7280',
          }}
        >
          This question will be removed from the questionnaire and the
          SharePoint list. Previously submitted responses will remain unchanged.
        </div> */}
      </DialogContent>

      <DialogActions
        sx={{
          px: 4,
          py: 2.5,
          gap: 1.5,
          backgroundColor: '#ffffff',
        }}
      >
        <Button
          onClick={() => setDeleteTarget(null)}
          sx={{
            color: '#004632',
            fontWeight: 700,
            textTransform: 'uppercase',
            px: 2.5,
          }}
        >
          Cancel
        </Button>

        <Button
          variant="contained"
          onClick={() => void handleConfirmDeleteQuestion()}
          sx={{
            backgroundColor: '#dc2626',
            color: '#ffffff',
            fontWeight: 700,
            textTransform: 'uppercase',
            borderRadius: '7px',
            px: 3,
            py: 1,
            boxShadow: '0 4px 10px rgba(220, 38, 38, 0.25)',
            '&:hover': {
              backgroundColor: '#b91c1c',
            },
          }}
        >
          Remove Question
        </Button>
      </DialogActions>
    </Dialog>
  );
};

  const selectedLabel = MEETING_OPTIONS.find(m => m.value === selectedMeeting)?.label ?? '';

  const filteredRows = useMemo(
    () => questions.filter(q => {
      const meetingMatch = q.meetingNo === selectedMeeting;
      const statusMatch = selectedStatus === 'All' || q.status === selectedStatus;
      return meetingMatch && statusMatch;
    }),
    [questions, selectedMeeting, selectedStatus],
  );

  const visibleRows = useMemo(() => {
    const term = searchText.trim().toLowerCase();
    const matchingRows = term ? filteredRows.filter(r => r.text.toLowerCase().includes(term)) : filteredRows;
    return [...matchingRows].sort((left, right) => {
      const result = left.spId - right.spId;
      return questionSortDirection === 'asc' ? result : -result;
    });
  }, [filteredRows, searchText, questionSortDirection]);

  const activeForMeeting = filteredRows.filter(q => q.status === 'Active').length;

  const meetingSummary = useMemo(
    () => MEETING_OPTIONS.map(m => ({
      ...m,
      total:  questions.filter(q => q.meetingNo === m.value).length,
      active: questions.filter(q => q.meetingNo === m.value && q.status === 'Active').length,
    })),
    [questions],
  );

  const handleMeetingChange = (val: number) => {
    setSelectedMeeting(val);
    setSearchText('');
    setDrawerDupWarning('');
  };

  const handleStatusChange = (status: StatusFilter) => {
    setSelectedStatus(status);
    setSearchText('');
  };

  const drawerMeetingLabel = MEETING_OPTIONS.find(m => m.value === drawerMeeting)?.label ?? '';


  const columns: GridColDef[] = [
    {
      field: 'serial',
      headerName: '#',
      width: 52,
      sortable: false,
      renderCell: (p: GridRenderCellParams) => (
        <span style={{ color: '#6b7280', fontSize: '0.75rem', fontWeight: 600 }}>
          {visibleRows.findIndex(r => r.id === p.row.id) + 1}
        </span>
      ),
    },
    {
      field: 'text',
      headerName: 'Question',
      flex: 1,
      minWidth: 300,
      sortable: false,
      renderCell: (p: GridRenderCellParams) => (
        <span style={{
          fontSize: '0.813rem',
          color: p.row.status === 'Active' ? '#111827' : '#9ca3af',
          textDecoration: p.row.status === 'Active' ? 'none' : 'line-through',
          lineHeight: 1.5,
          whiteSpace: 'normal',
          padding: '0.5rem 0',
        }}>
          {p.value as string}
        </span>
      ),
    },
    {
      field: 'status',
      headerName: 'Status',
      width: 115,
      sortable: false,
      renderCell: (p: GridRenderCellParams) => (
        <Chip
          label={p.value as string}
          size="small"
          color={p.value === 'Active' ? 'success' : 'default'}
          sx={{ fontSize: '0.7rem', height: '22px', fontWeight: 600 }}
        />
      ),
    },
    {
      field: 'toggle',
      headerName: 'Enable / Disable',
      width: 150,
      sortable: false,
      renderCell: (p: GridRenderCellParams) => {
        const q      = p.row as Question;
        const saving = (rowSaveStates[q.id] ?? 'idle') === 'saving';
        const isOn   = q.status === 'Active';

        return (
          <div style={{ display: 'flex', alignItems: 'center', gap: '0.375rem' }}>
            {saving && (
              <Tooltip title="Saving…">
                <SyncIcon sx={{ fontSize: '0.85rem', color: '#f59e0b', animation: 'spin 1s linear infinite' }} />
              </Tooltip>
            )}
            <span style={{
              fontSize: '0.7rem',
              color: isOn ? '#6b7280' : '#78736e',
              fontWeight: 500,
              minWidth: 22,
            }}>
              {isOn ? 'On' : 'Off'}
            </span>
            <Switch
              checked={isOn}
              onChange={() => void handleToggleStatus(q)}
              disabled={saving}
              size="small"
              sx={{
                '& .MuiSwitch-switchBase.Mui-checked':                    { color: '#00D264' },
                '& .MuiSwitch-switchBase.Mui-checked + .MuiSwitch-track': { backgroundColor: '#00D264' },
              }}
            />
          </div>
        );
      },
    },
    {
      field: 'actions',
      headerName: 'Actions',
      width: 120,
      sortable: false,
      filterable: false,
      disableColumnMenu: true,
      align: 'center',
      headerAlign: 'center',
      renderCell: (p: GridRenderCellParams) => {
        const question = p.row as Question;
        const saving = (rowSaveStates[question.id] ?? 'idle') === 'saving';
        return (
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '0.25rem' }}>
            <Tooltip title="Update question" arrow>
              <span>
                <IconButton
                  className={styles.editActionBtn}
                  size="small"
                  onClick={() => openEditQuestionDrawer(question)}
                  disabled={saving || isAdding}
                  aria-label={`Update question: ${question.text}`}
                >
                  <EditOutlinedIcon sx={{ fontSize: '1rem' }} />
                </IconButton>
              </span>
            </Tooltip>
            <Tooltip title="Delete question" arrow>
              <span>
                <IconButton
                  size="small"
                  onClick={() => void handleDeleteQuestion(question)}
                  disabled={saving || isAdding}
                  aria-label={`Delete question: ${question.text}`}
                  sx={{ color: '#dc2626', '&:hover': { backgroundColor: '#fee2e2' } }}
                >
                  <DeleteOutlineIcon sx={{ fontSize: '1rem' }} />
                </IconButton>
              </span>
            </Tooltip>
          </div>
        );
      },
    },
  ];


  return (
    <div className={styles.Container}>
      <BuddyLoader
        visible={loaderVisible}
        variant="default"
        message={loaderMessage}
      />

      <div className={styles.pageHeading}>
        <div className={styles.pageHeadingLeft}>
          <div className={styles.pageHeadingIcon}>
            <QuizOutlinedIcon sx={{ fontSize: '1.25rem' }} />
          </div>
          <div>
            <h2 className={styles.pageTitle}>Buddy Questionnaire</h2>
            <p className={styles.pageSubtitle}>Manage questions for each buddy meeting</p>
          </div>
        </div>

        <button
          className={styles.openAddDrawerBtn}
          onClick={openAddQuestionDrawer}
        >
          <AddCircleOutlineIcon sx={{ fontSize: '1rem' }} />
          Add New Question
        </button>
      </div>

      <div className={styles.meetingHeader}>
        {/* <div className={styles.meetingHeaderLeft}>
          <div className={styles.meetingHeaderIcon}>
            <span>📊</span>
          </div>
          <div>
            <h2 className={styles.meetingHeaderTitle}>Meeting Questions</h2>
            <p className={styles.meetingHeaderSubtitle}>
              {activeForMeeting} active of {filteredRows.length} total questions
            </p>
          </div>
        </div> */}

        <div className={styles.meetingChipsContainer}>
          {meetingSummary.map(m => (
            <button
              key={m.value}
              className={`${styles.meetingChip} ${selectedMeeting === m.value ? styles.meetingChipActive : ''}`}
              onClick={() => handleMeetingChange(m.value)}
            >
              <span className={styles.meetingChipLabel}>{m.label}</span>
              <span className={styles.meetingChipCount}>
                <span className={styles.meetingChipActiveCount}>{m.active}</span>
                <span className={styles.meetingChipSep}>/</span>
                <span>{m.total}</span>
              </span>
            </button>
          ))}
        </div>
      </div>

      {loadError && (
        <div className={styles.errorBanner}>
          <ErrorOutlineIcon sx={{ fontSize: '1rem' }} />
          {loadError}
          <button className={styles.retryBtn} onClick={() => void loadQuestions()}>
            Retry
          </button>
        </div>
      )}

      <div className={styles.tableSection}>
        <div className={styles.tableHeader}>
          <div className={styles.tableHeaderLeft}>
            <h3 className={styles.tableTitle}>
              Questions — <span className={styles.tableTitleMeeting}>{selectedLabel}</span>
            </h3>
            <div className={styles.tableStats}>
              <span className={styles.tableStatActive}>{activeForMeeting} active</span>
              <span className={styles.tableStatDivider}>/</span>
              <span className={styles.tableStatTotal}>{filteredRows.length} total</span>
              {searchText.trim() && (
                <span className={styles.tableStatSearch}>
                  &nbsp;· {visibleRows.length} match
                </span>
              )}
            </div>
          </div>

          <div className={styles.tableHeaderRight}>
            <IconButton
              size="small"
              onClick={() => setQuestionSortDirection((direction) => direction === 'asc' ? 'desc' : 'asc')}
              aria-label={`Sort questions ${questionSortDirection === 'asc' ? 'descending' : 'ascending'}`}
              title={`Sort ${questionSortDirection === 'asc' ? 'descending' : 'ascending'}`}
            >
              {questionSortDirection === 'asc' ? <ArrowUpwardIcon sx={{ fontSize: '1rem' }} /> : <ArrowDownwardIcon sx={{ fontSize: '1rem' }} />}
            </IconButton>
            <div className={styles.searchWrap}>
              <SearchIcon sx={{ fontSize: '0.9rem', color: '#004632', flexShrink: 0 }} />
              <input
                type="text"
                className={styles.searchInput}
                placeholder="Search questions…"
                value={searchText}
                onChange={e => setSearchText(e.target.value)}
              />
              {searchText && (
                <button
                  className={styles.searchClear}
                  onClick={() => setSearchText('')}
                  aria-label="Clear search"
                >
                  <CloseIcon sx={{ fontSize: '0.75rem' }} />
                </button>
              )}
            </div>

            <div className={styles.statusLegend}>
              {/* <span className={styles.legendItem}>
                <span className={styles.legendDotGreen} /> Active
              </span>
              <span className={styles.legendItem}>
                <span className={styles.legendDotGray} /> InActive
              </span> */}
              <RadioGroup
            row
            value={selectedStatus}
            onChange={(e) => handleStatusChange(e.target.value as StatusFilter)}
            sx={{
              display: 'flex',
              // gap: '2rem',
            }}
          >
            <FormControlLabel
              value="All"
              control={<Radio size="small" sx={{
                color: '#6b7280',
                '&.Mui-checked': { color: '#004632' },
              }} />}
              label={<span className={styles.legendItem}>All</span>}
            />
            <FormControlLabel
              value="Active"
              control={<Radio size="small" sx={{
                color: '#6b7280',
                '&.Mui-checked': { color: '#00D264' },
              }} />}
              label={<span className={styles.legendItem}>Active</span>}
            />
            <FormControlLabel
              value="InActive"
              control={<Radio size="small" sx={{
                color: '#6b7280',
                '&.Mui-checked': { color: '#ef4444' },
              }} />}
              label={<span className={styles.legendItem}>InActive</span>}
            />
          </RadioGroup>
            </div>
          </div>
        </div>

        <div className={styles.gridCard}>
          {isLoading ? (
            <div className={styles.loadingState}>
              <SyncIcon sx={{ fontSize: '2rem', color: '#00D264', animation: 'spin 1s linear infinite' }} />
              <p>Loading questions…</p>
            </div>
          ) : filteredRows.length === 0 ? (
            <div className={styles.emptyState}>
              <QuizOutlinedIcon sx={{ fontSize: '2.5rem', color: '#d1d5db' }} />
              <p>No questions for {selectedLabel}.</p>
              {/* <p>
                <button className={styles.openAddDrawerBtn} style={{ fontSize: '0.75rem', padding: '0.4rem 0.85rem' }} onClick={() => setDrawerOpen(true)}>
                  <AddCircleOutlineIcon sx={{ fontSize: '0.85rem' }} /> Add the first question
                </button>
              </p> */}
            </div>
          ) : visibleRows.length === 0 ? (
            <div className={styles.emptyState}>
              <SearchIcon sx={{ fontSize: '2.5rem', color: '#d1d5db' }} />
              <p>No questions match &ldquo;{searchText}&rdquo;</p>
              <p>
                Try a different term or{' '}
                <button className={styles.clearSearchLink} onClick={() => setSearchText('')}>
                  clear the search
                </button>.
              </p>
            </div>
          ) : (
            <DataGrid
              rows={visibleRows}
              columns={columns}
              disableRowSelectionOnClick
              hideFooterSelectedRowCount
              initialState={{ pagination: { paginationModel: { pageSize: 10 } } }}
              pageSizeOptions={[5, 10, 20]}
              rowHeight={56}
              columnHeaderHeight={40}
              getRowClassName={p => p.row.status !== 'Active' ? 'row-inactive' : ''}
              sx={{
                border: 'none',
                width: '100%',
                height: '100%',
                '& .MuiDataGrid-columnHeaders': {
                  backgroundColor: '#004632',
                  color: '#fff',
                  fontSize: '0.72rem',
                  fontWeight: 700,
                  minHeight: '40px !important',
                  maxHeight: '40px !important',
                },
                '& .MuiDataGrid-columnHeaderTitle': {
                  fontWeight: 700,
                  textTransform: 'uppercase',
                  letterSpacing: '0.05em',
                  color: '#fff',
                },
                '& .MuiDataGrid-sortIcon':        { color: '#fff !important' },
                '& .MuiDataGrid-menuIconButton':  { color: '#fff !important' },
                '& .MuiDataGrid-columnSeparator': { color: 'rgba(255,255,255,0.15)' },
                '& .MuiDataGrid-row:hover': { backgroundColor: '#e6fff0 !important' },
                '& .row-inactive': {
                  backgroundColor: '#f9fafb',
                  opacity: 0.72,
                  '&:hover': { opacity: 1 },
                },
                '& .MuiDataGrid-cell': {
                  fontSize: '0.813rem',
                  padding: '0 0.75rem',
                  borderColor: '#f3f4f6',
                  alignItems: 'center',
                },
                '& .MuiDataGrid-footerContainer': {
                  borderTop: '1px solid #e5e7eb',
                  minHeight: '42px',
                  maxHeight: '42px',
                },
                '& .MuiTablePagination-root': { fontSize: '0.75rem' },
                '& .MuiTablePagination-selectLabel, & .MuiTablePagination-displayedRows': {
                  fontSize: '0.725rem',
                },
              }}
            />
          )}
        </div>
      </div>

      <Drawer
        anchor="right"
        open={drawerOpen}
        onClose={handleDrawerClose}
        sx={{
          ...DRAWER_OVER_TOPBAR_SX,
          '& .MuiDrawer-paper': {
            ...DRAWER_OVER_TOPBAR_SX['& .MuiDrawer-paper'],
            width: { xs: '100%', sm: 420 },
            boxShadow: '-8px 0 40px rgba(0,0,0,0.18)',
          },
        }}
      >
        <div className={styles.addQuestionDrawer}>

          <div className={styles.addDrawerHeader}>
            <div style={{ flex: 1, minWidth: 0 }}>
              <h3 className={styles.addDrawerTitle}>
                {editingQuestion ? 'Update Question' : 'Add New Question'}
              </h3>
              <p className={styles.addDrawerSub}>
                Question will be {editingQuestion ? 'updated in' : 'added to'}&nbsp;
                <strong style={{ color: '#8CFF8C' }}>{drawerMeetingLabel}</strong>
              </p>
            </div>
            <IconButton
              className={styles.addDrawerCloseBtn}
              onClick={handleDrawerClose}
              disabled={isAdding}
              size="small"
              aria-label="Close drawer"
            >
              <CloseIcon sx={{ fontSize: '1.1rem' }} />
            </IconButton>
          </div>

          <div style={{
            height: 3,
            background: 'linear-gradient(90deg, #00D264, #8CFF8C, #004632)',
            flexShrink: 0,
          }} />

          <div className={styles.addDrawerBody}>

            {drawerSuccess ? (
              <div className={styles.drawerSuccessState}>
                <div className={styles.drawerSuccessIcon}>
                  <CheckCircleOutlineIcon sx={{ fontSize: '3rem', color: '#00D264' }} />
                </div>
                <div className={styles.drawerSuccessTitle}>
                  {editingQuestion ? 'Question Updated!' : 'Question Added!'}
                </div>
                <div className={styles.drawerSuccessMsg}>
                  Successfully {editingQuestion ? 'updated in' : 'added to'} <strong>{drawerMeetingLabel}</strong>
                </div>
              </div>
            ) : (
              <>
                <div className={styles.drawerField}>
                  <label className={styles.drawerFieldLabel}>
                    Select Meeting
                  </label>
                  <Select
                    value={String(drawerMeeting)}
                    onChange={(e: SelectChangeEvent) => {
                      setDrawerMeeting(Number(e.target.value));
                      setDrawerDupWarning('');
                    }}
                    fullWidth
                    size="small"
                    IconComponent={KeyboardArrowDownIcon}
                    MenuProps={{
                      sx: {
                        zIndex: '100001 !important',
                      },
                    }}
                    sx={{
                      fontSize: '0.813rem',
                      background: '#fff',
                      '& .MuiOutlinedInput-notchedOutline':             { borderColor: '#e5e7eb' },
                      '&:hover .MuiOutlinedInput-notchedOutline':       { borderColor: '#00D264' },
                      '&.Mui-focused .MuiOutlinedInput-notchedOutline': { borderColor: '#00D264' },
                    }}
                  >
                    {MEETING_OPTIONS.map(m => (
                      <MenuItem key={m.value} value={String(m.value)} sx={{ fontSize: '0.813rem' }}>
                        {m.label}
                      </MenuItem>
                    ))}
                  </Select>

                  <div className={styles.drawerMeetingStats}>
                    {(() => {
                      const summary = meetingSummary.find(m => m.value === drawerMeeting);
                      return summary ? (
                        <>
                          <span className={styles.drawerStatActive}>{summary.active} active</span>
                          <span className={styles.drawerStatSep}>/</span>
                          <span className={styles.drawerStatTotal}>{summary.total} total</span>
                        </>
                      ) : null;
                    })()}
                  </div>
                </div>

                <div className={styles.drawerField} style={{ flex: 1 }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                    <label className={styles.drawerFieldLabel}>
                      Question <span className={styles.required}>*</span>
                    </label>
                    <span className={`${styles.drawerCharCount} ${drawerCharCount > 220 ? styles.drawerCharCountWarn : ''}`}>
                      {drawerCharCount}/255
                    </span>
                  </div>
                  <textarea
                    ref={textareaRef}
                    className={`${styles.drawerTextarea} ${drawerDupWarning ? styles.textareaError : ''}`}
                    placeholder={`e.g. "How are you settling into the team so far?"`}
                    value={drawerText}
                    maxLength={255}
                    onChange={e => {
                      setDrawerText(e.target.value);
                      setDrawerCharCount(e.target.value.length);
                      if (drawerDupWarning) setDrawerDupWarning('');
                    }}
                    onKeyDown={e => {
                      if (e.key === 'Enter' && !e.shiftKey) {
                        e.preventDefault();
                        void handleSaveQuestion();
                      }
                    }}
                    rows={5}
                    disabled={isAdding}
                  />
                  {/* <p className={styles.drawerHint}>
                    Press <kbd className={styles.drawerKbd}>Enter</kbd> to submit &nbsp;·&nbsp;
                    <kbd className={styles.drawerKbd}>Shift + Enter</kbd> for new line
                  </p> */}
                </div>

                {drawerDupWarning && (
                  <div className={styles.duplicateWarning}>
                    <WarningAmberIcon sx={{ fontSize: '0.9rem', flexShrink: 0 }} />
                    <span>{drawerDupWarning}</span>
                    <button
                      className={styles.warnDismiss}
                      onClick={() => setDrawerDupWarning('')}
                      aria-label="Dismiss warning"
                    >
                      <CloseIcon sx={{ fontSize: '0.8rem' }} />
                    </button>
                  </div>
                )}

                {/*
                {drawerText.trim() && (
                  <div className={styles.drawerPreviewCard}>
                    <div className={styles.drawerPreviewLabel}>Preview</div>
                    <div className={styles.drawerPreviewText}>{drawerText.trim()}</div>
                    <div className={styles.drawerPreviewMeta}>
                      <span className={styles.drawerPreviewBadge}>{drawerMeetingLabel}</span>
                      <span className={styles.drawerPreviewStatus}>● Active</span>
                    </div>
                  </div>
                )} */}
              </>
            )}
          </div>

          {!drawerSuccess && (
            <div className={styles.addDrawerFooter}>
              <button
                className={styles.cancelBtn}
                onClick={handleDrawerClose}
                disabled={isAdding}
              >
                Cancel
              </button>
              <button
                className={styles.addBtn}
                onClick={() => void handleSaveQuestion()}
                disabled={!drawerText.trim() || isAdding}
              >
                {isAdding
                  ? <SyncIcon sx={{ fontSize: '1rem', animation: 'spin 1s linear infinite' }} />
                  : editingQuestion
                    ? <EditOutlinedIcon sx={{ fontSize: '1rem' }} />
                    : <AddCircleOutlineIcon sx={{ fontSize: '1rem' }} />
                }
                {isAdding
                  ? editingQuestion ? 'Updating…' : 'Adding…'
                  : editingQuestion ? 'Update Question' : 'Add Question'}
              </button>
            </div>
          )}
        </div>
      </Drawer>
      {/* Delete Confirmation Dialog */}
      {renderDeleteConfirmation()}

      {/* <Snackbar
        open={Boolean(deleteTarget)}
        anchorOrigin={{ vertical: 'bottom', horizontal: 'right' }}
        onClose={() => setDeleteTarget(null)}
        message={
          <div style={{ maxWidth: 420 }}>
            <div style={{ fontWeight: 700, marginBottom: 4 }}>Remove this question?</div>
            <div style={{ fontSize: '0.8rem', lineHeight: 1.45 }}>
              It will no longer appear in {deleteTarget?.meeting || 'this meeting'} questionnaire. Previously submitted responses will remain unchanged.
            </div>
          </div>
        }
        action={
          <div style={{ display: 'flex', gap: 6 }}>
            <Button size="small" onClick={() => setDeleteTarget(null)} sx={{ color: '#ffffff', textTransform: 'none' }}>
              Keep
            </Button>
            <Button
              size="small"
              onClick={() => void handleConfirmDeleteQuestion()}
              sx={{ color: '#fecaca', fontWeight: 700, textTransform: 'none' }}
            >
              Remove
            </Button>
          </div>
        }
        ContentProps={{
          sx: {
            backgroundColor: '#1f2937',
            borderRadius: '10px',
            boxShadow: '0 10px 30px rgba(0,0,0,0.22)',
            alignItems: 'center',
            minWidth: { xs: 'calc(100vw - 32px)', sm: '460px' },
          },
        }}
      /> */}
    </div>
  );
};

export default BuddyQuestionnaire;