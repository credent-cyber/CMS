/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable  @typescript-eslint/no-explicit-any*/
/* eslint-disable no-void*/
import * as React from 'react';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { DataGrid, GridColDef, GridRenderCellParams } from '@mui/x-data-grid';
import { Chip, CircularProgress, Drawer, IconButton, RadioGroup, FormControlLabel, Radio, Switch, Button ,
   Dialog,
    DialogTitle,
    DialogContent,
    DialogActions,
} from '@mui/material';
import ClearIcon from '@mui/icons-material/Clear';
import CloseIcon from '@mui/icons-material/Close';
import HelpOutlineIcon from '@mui/icons-material/HelpOutline';
import EditOutlinedIcon from '@mui/icons-material/EditOutlined';
import DeleteOutlineIcon from '@mui/icons-material/DeleteOutline';
import ArrowUpwardIcon from '@mui/icons-material/ArrowUpward';
import ArrowDownwardIcon from '@mui/icons-material/ArrowDownward';
import { sp } from '@pnp/sp';
import { useSnackbar } from './Snackbar';
import { BuddyLoader } from './Buddyloader';
import LIST_CONFIG from '../../../config/spListConfig';
import styles from './JoineeQuestionnaire.module.scss';
import type { IBuddyAppProps } from './IBuddyAppProps';


interface Question {
  id: string;           
  spId: number;         
  title: string;        
  questionId: string;   
  milestone: string;   
  questionType: string; 
  text: string;         
  options: string;      
  isRequired: string;   
  sortOrder: number;    
  status: SPStatus;    
}

type SPStatus = 'Active' | 'InActive';

type RowSaveState = 'idle' | 'saving';

const BRAND = {
  darkGreen:    '#004632', 
  classicGreen: '#00D264', 
  lightGreen:   '#8CFF8C', 
} as const;

const SWITCH_SX = {
  '& .MuiSwitch-switchBase.Mui-checked': { color: '#00D264' },
  '& .MuiSwitch-switchBase.Mui-checked + .MuiSwitch-track': {
    backgroundColor: '#00D264',
  },
} as const;

const QUESTION_TYPES = {
  yes_no: {
    label: 'Yes/No',
    options: ['Yes', 'No'],
  },
  agree_disagree: {
    label: 'Agree/Disagree',
    options: ['Strongly Agree', 'Agree', 'Neutral', 'Disagree', 'Strongly Disagree'],
  },
  rating: {
    label: 'Rating',
    options: ['1', '2', '3', '4', '5'],
  },
  remark: {
    label: 'Remark/Comment',
    options: [],
  },
};

const normalizeQuestionTypeKey = (value: string): keyof typeof QUESTION_TYPES => {
  const normalized = String(value || '').trim().toLowerCase().replace(/[\s/-]+/g, '_');
  if (normalized === 'yesno' || normalized === 'yes_no') return 'yes_no';
  if (normalized === 'likert' || normalized === 'agree_disagree') return 'agree_disagree';
  if (normalized === 'rating') return 'rating';
  return 'remark';
};

const getMilestoneCode = (milestone: string): string => {
  return milestone === '3rd' ? '3m' : '6m';
};

const normalizeSingleSpaces = (value: string): string =>
  String(value || '').trim().replace(/\s+/g, ' ');

const normalizeQuestionText = (value: string): string =>
  normalizeSingleSpaces(value).toLowerCase();

const getStatusLabel = (status: SPStatus): string =>
  status === 'Active' ? 'Active' : 'Deactive';

const DRAWER_OVER_TOPBAR_SX = {
  zIndex: 99999,
  "& .MuiBackdrop-root": {
    zIndex: 99998,
  },
  "& .MuiDrawer-paper": {
    top: "0 !important",
    height: "100vh !important",
    maxHeight: "100vh !important",
    zIndex: 100000,
  },
} as const;

const getNextQuestionSequence = (milestoneQuestions: Question[]): number => {
  const usedNumbers = milestoneQuestions
    .map((q) => {
      const match = String(q.questionId || '').match(/_q(\d+)$/i);
      return match ? Number(match[1]) : Number(q.sortOrder || 0);
    })
    .filter((n) => Number.isFinite(n) && n > 0);

  return usedNumbers.length > 0 ? Math.max(...usedNumbers) + 1 : 1;
};


export const JoineeQuestionnaire: React.FC<IBuddyAppProps> = (props) => {
  const containerRef       = useRef<HTMLDivElement>(null);
  const isMountedRef       = useRef(true);
  const { showSnackbar }   = useSnackbar();

  const [questions,        setQuestions]        = useState<Question[]>([]);
  const [isLoading,        setIsLoading]        = useState(false);
  const [loadError,        setLoadError]        = useState('');

  const [newQuestionText,  setNewQuestionText]  = useState<string>('');
  const [isAdding,         setIsAdding]         = useState<boolean>(false);
  const [duplicateWarning, setDuplicateWarning] = useState<string>('');
  const [searchText,       setSearchText]       = useState<string>('');
  const [selectedMilestone, setSelectedMilestone] = useState<string>('3rd');
  const [selectedQuestionType, setSelectedQuestionType] = useState<string>('yes_no');
  const [isRequired,       setIsRequired]       = useState<string>('No');
  const [isActive,         setIsActive]         = useState<boolean>(true);
  const [isAddDrawerOpen,  setIsAddDrawerOpen]  = useState<boolean>(false);
  const [editingQuestion,  setEditingQuestion]  = useState<Question | null>(null);
  const [deleteTarget,     setDeleteTarget]     = useState<Question | null>(null);
  const [selectedStatus,   setSelectedStatus]   = useState<'All' | 'Active' | 'InActive'>('All');
  const [questionSortDirection, setQuestionSortDirection] = useState<'asc' | 'desc'>('asc');

  const [rowSaveStates, setRowSaveStates] = useState<Record<string, RowSaveState>>({});
  const [loaderVisible,    setLoaderVisible]    = useState<boolean>(false);
  const [loaderMessage,    setLoaderMessage]    = useState<string>('Processing your request…');

  useEffect(() => { return () => { isMountedRef.current = false; }; }, []);

  const loadQuestions = useCallback(async () => {
    if (!props.context) return;
    setIsLoading(true);
    setLoadError('');
    setLoaderVisible(true);
    setLoaderMessage('Loading questions…');
    try {
      const items = await sp.web.lists
        .getByTitle(LIST_CONFIG.LISTS.FEEDBACK_QUESTIONS_MASTER)
        .items
        .select('ID,Title,QuestionId,Milestone,QuestionType,QuestionText,Options,IsRequired,SortOrder,Status')
        .orderBy('SortOrder')
        .top(5000)
        .get();

      if (!isMountedRef.current) return;

      setQuestions(
        (items || []).map((i: any) => ({
          id:           i.ID.toString(),
          spId:         i.ID,
          title:        i.Title || '',
          questionId:   i.QuestionId || '',
          milestone:    i.Milestone || '',
          questionType: normalizeQuestionTypeKey(i.QuestionType || 'remark'),
          text:         i.QuestionText || '',
          options:      i.Options || '',
          isRequired:   i.IsRequired || 'No',
          sortOrder:    i.SortOrder || i.ID,
          status:       (i.Status === 'Active' ? 'Active' : 'InActive') as SPStatus,
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

  const resetAddQuestionForm = () => {
    setNewQuestionText('');
    setDuplicateWarning('');
    setIsRequired('No');
    setSelectedQuestionType('yes_no');
    setIsActive(true);
    setEditingQuestion(null);
  };

  const openAddQuestionDrawer = () => {
    resetAddQuestionForm();
    setIsAddDrawerOpen(true);
  };

  const openEditQuestionDrawer = (question: Question) => {
    setEditingQuestion(question);
    setSelectedMilestone(question.milestone);
    setSelectedQuestionType(normalizeQuestionTypeKey(question.questionType || 'yes_no'));
    setIsRequired(question.isRequired || 'No');
    // Preserve the question's current status while editing.
    setIsActive(question.status === 'Active');
    setNewQuestionText(question.text);
    setDuplicateWarning('');
    setIsAddDrawerOpen(true);
  };

  const closeAddQuestionDrawer = () => {
    if (isAdding) return;
    setIsAddDrawerOpen(false);
    resetAddQuestionForm();
  };

  const preserveSubmittedFeedbackQuestionSnapshot = async (question: Question) => {
    const questionId = String(question.questionId || '').trim();
    if (!questionId) return;

    const allocationList = sp.web.lists.getByTitle(LIST_CONFIG.LISTS.BuddyAllocate);

    const preserveForItems = async (items: any[]) => {
      for (const item of items || []) {
        const rawHistory = String(item.FeedbackJsonHistory || '').trim();
        if (!rawHistory || !rawHistory.includes(questionId)) continue;

        let entries: any[] = [];
        try {
          const parsed = JSON.parse(rawHistory);
          entries = Array.isArray(parsed) ? parsed : [parsed];
        } catch {
          // Ignore malformed/legacy history instead of blocking unrelated records.
          continue;
        }

        let changed = false;
        const patchedEntries = entries.map((entry: any) => {
          const answers = entry?.answers && typeof entry.answers === 'object'
            ? entry.answers
            : {};

          if (!Object.prototype.hasOwnProperty.call(answers, questionId)) return entry;

          const existingSnapshots = entry?.questionSnapshots && typeof entry.questionSnapshots === 'object'
            ? entry.questionSnapshots
            : {};

          // Never overwrite a snapshot that was already captured at submission time.
          if (existingSnapshots[questionId]?.question) return entry;

          changed = true;
          return {
            ...entry,
            questionSnapshots: {
              ...existingSnapshots,
              [questionId]: {
                question: normalizeSingleSpaces(question.text),
                type: normalizeQuestionTypeKey(question.questionType || 'remark'),
              },
            },
          };
        });

        if (changed) {
          await allocationList.items
            .getById(Number(item.ID))
            .update({ FeedbackJsonHistory: JSON.stringify(patchedEntries, null, 2) });
        }
      }
    };

    let page = await allocationList.items
      .select('ID,FeedbackJsonHistory')
      .top(500)
      .getPaged();

    await preserveForItems(page.results);

    while (page.hasNext) {
      page = await page.getNext();
      await preserveForItems(page.results);
    }
  };

  const handleSaveQuestion = async () => {
    const trimmed = normalizeSingleSpaces(newQuestionText);
    if (!trimmed || isAdding) return;

    const normalizedText = normalizeQuestionText(trimmed);
    const activeDuplicate = questions.find(
      q => q.id !== editingQuestion?.id
        && q.status === 'Active'
        && q.milestone === selectedMilestone
        && normalizeQuestionText(q.text) === normalizedText,
    );

    if (isActive && activeDuplicate) {
      setDuplicateWarning(
        `This question is already Active in ${selectedMilestone} Meeting. Please deactivate the existing question before adding it again.`,
      );
      return;
    }

    setDuplicateWarning('');
    setIsAdding(true);
    setLoaderVisible(!editingQuestion);
    if (!editingQuestion) setLoaderMessage('Adding new question to milestone…');

    try {
      const selectedQType = QUESTION_TYPES[selectedQuestionType as keyof typeof QUESTION_TYPES];
      const optionsString = selectedQType?.options.length 
        ? selectedQType.options.join('|')
        : '';

      if (editingQuestion) {
        // Update the existing SharePoint item in place. Do not create a new
        // question/version when editing, and preserve the current status.
        const nextStatus = (isActive ? 'Active' : 'InActive') as SPStatus;
        const questionList = sp.web.lists.getByTitle(
          LIST_CONFIG.LISTS.FEEDBACK_QUESTIONS_MASTER,
        );

        // Freeze the old question text/type inside every already-submitted
        // feedback entry before changing the master question. This prevents
        // historical submissions from showing the newly edited wording.
        setLoaderVisible(true);
        setLoaderMessage('Preserving submitted feedback history…');
        await preserveSubmittedFeedbackQuestionSnapshot(editingQuestion);

        await questionList.items
          .getById(editingQuestion.spId)
          .update({
            Title:        trimmed.substring(0, 255),
            QuestionText: trimmed,
            QuestionType: selectedQuestionType,
            Options:      optionsString,
            IsRequired:   isRequired,
            Status:       nextStatus,
          });

        if (!isMountedRef.current) return;

        setQuestions(prev => prev.map(item =>
          item.id === editingQuestion.id
            ? {
                ...item,
                title:        trimmed.substring(0, 255),
                questionType: selectedQuestionType,
                text:         trimmed,
                options:      optionsString,
                isRequired,
                status:       nextStatus,
              }
            : item,
        ));
        setSearchText('');
        setIsAddDrawerOpen(false);
        setEditingQuestion(null);
        showSnackbar(
          `Question updated successfully and remains ${nextStatus === 'Active' ? 'Active' : 'Inactive'}.`,
          'success',
          3000,
        );
        return;
      }

      const milestoneQuestions = questions.filter(q => q.milestone === selectedMilestone);
      const newSortOrder = getNextQuestionSequence(milestoneQuestions);

      const milestoneCode = getMilestoneCode(selectedMilestone);
      const generatedQuestionId = `${milestoneCode}_q${newSortOrder}`;

      const result = await sp.web.lists
        .getByTitle(LIST_CONFIG.LISTS.FEEDBACK_QUESTIONS_MASTER)
        .items.add({
          Title:        trimmed.substring(0, 255),
          QuestionText: trimmed,
          QuestionId:   generatedQuestionId,
          Milestone:    selectedMilestone,
          QuestionType: selectedQuestionType,
          Options:      optionsString,
          IsRequired:   isRequired,
          SortOrder:    newSortOrder,
          Status:       isActive ? 'Active' : 'InActive',
        });

      if (!isMountedRef.current) return;

      const newId = result.data.ID;
      setQuestions(prev => [...prev, {
        id:           newId.toString(),
        spId:         newId,
        title:        trimmed.substring(0, 255),
        questionId:   generatedQuestionId,
        milestone:    selectedMilestone,
        questionType: selectedQuestionType,
        text:         trimmed,
        options:      optionsString,
        isRequired:   isRequired,
        sortOrder:    newSortOrder,
        status:       (isActive ? 'Active' : 'InActive') as SPStatus,
      }]);
      setNewQuestionText('');
      setSearchText('');
      setIsRequired('No');
      setIsActive(true);
      setIsAddDrawerOpen(false);

      showSnackbar(
        `Question "${generatedQuestionId}" added successfully to ${selectedMilestone} Milestone!`,
        'success',
        3000,
      );
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

  const handleToggleStatus = async (q: Question) => {
    const nextStatus: SPStatus = q.status === 'Active' ? 'InActive' : 'Active';

    if (nextStatus === 'Active') {
      const normalizedText = normalizeQuestionText(q.text);
      const activeDuplicate = questions.find(
        item => item.id !== q.id
          && item.status === 'Active'
          && item.milestone === q.milestone
          && normalizeQuestionText(item.text) === normalizedText,
      );

      if (activeDuplicate) {
        showSnackbar(
          `Cannot activate this question because the same question is already Active in ${q.milestone} Meeting.`,
          'error',
          4000,
        );
        return;
      }
    }

    setQuestions(prev => prev.map(item =>
      item.id === q.id ? { ...item, status: nextStatus } : item,
    ));
    setRowSaveStates(prev => ({ ...prev, [q.id]: 'saving' }));
    setLoaderVisible(true);
    setLoaderMessage('Updating question status…');

    try {
      await sp.web.lists
        .getByTitle(LIST_CONFIG.LISTS.FEEDBACK_QUESTIONS_MASTER)
        .items.getById(q.spId)
        .update({ Status: nextStatus });

      if (!isMountedRef.current) return;

      showSnackbar(
        nextStatus === 'Active'
          ? `Question marked as Active.`
          : `Question marked as Deactive.`,
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
        .getByTitle(LIST_CONFIG.LISTS.FEEDBACK_QUESTIONS_MASTER)
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

  const filteredRows = useMemo(
    () => questions.filter(q => {
      const milestoneMatch = q.milestone === selectedMilestone;
      const statusMatch = selectedStatus === 'All' || q.status === selectedStatus;
      return milestoneMatch && statusMatch;
    }),
    [questions, selectedMilestone, selectedStatus],
  );

  const visibleRows = useMemo(() => {
    const term = searchText.trim().toLowerCase();
    const matchingRows = term
      ? filteredRows.filter(r => r.text.toLowerCase().includes(term))
      : filteredRows;
    return [...matchingRows].sort((left, right) => {
      const result = left.sortOrder - right.sortOrder || left.spId - right.spId;
      return questionSortDirection === 'asc' ? result : -result;
    });
  }, [filteredRows, searchText, questionSortDirection]);

  const totalQuestions = filteredRows.length;
  const activeQuestions = filteredRows.filter(q => q.status === 'Active').length;

  const milestones = ['3rd', '6th'];
  const getMilestoneStats = (milestone: string) => {
    const qs = questions.filter(q => q.milestone === milestone);
    return {
      total: qs.length,
      active: qs.filter(q => q.status === 'Active').length,
    };
  };

  const columns: GridColDef[] = [
    {
      field: 'serial',
      headerName: '#',
      width: 50,
      sortable: false,
      renderCell: (p: GridRenderCellParams) => (
        <span style={{ color: '#6b7280', fontSize: '0.75rem', fontWeight: 600 }}>
          {visibleRows.findIndex((r: Question) => r.id === (p.row as Question).id) + 1}
        </span>
      ),
    },
    {
      field: 'text',
      headerName: 'Question',
      flex: 1,
      minWidth: 280,
      sortable: false,
      renderCell: (p: GridRenderCellParams) => {
        const row = p.row as Question;
        return (
          <div style={{ display: 'flex', flexDirection: 'column', gap: '0.25rem', lineHeight: 1.4 }}>
            <span style={{
              fontSize: '0.813rem',
              color: row.status === 'Active' ? '#111827' : '#9ca3af',
              textDecoration: row.status === 'Active' ? 'none' : 'line-through',
              fontWeight: 500,
            }}>
              {normalizeSingleSpaces(row.text)}
            </span>
          </div>
        );
      },
    },
    {
      field: 'milestone',
      headerName: 'Milestone',
      width: 100,
      sortable: false,
      renderCell: (p: GridRenderCellParams) => {
        const milestone = (p.row as Question).milestone || '—';
        return <span style={{ fontSize: '0.7rem', color: '#6b7280' }}>{milestone}</span>;
      },
    },
    {
      field: 'questionType',
      headerName: 'Type',
      width: 90,
      sortable: false,
      renderCell: (p: GridRenderCellParams) => {
        const qType = normalizeQuestionTypeKey((p.row as Question).questionType || 'remark');
        const qTypeLabel = QUESTION_TYPES[qType].label;
        const typeColors: Record<string, string> = {
          remark: '#3b82f6',
          yes_no: '#10b981',
          agree_disagree: '#f59e0b',
          rating: '#8b5cf6',
        };
        return (
          <span style={{
            fontSize: '0.65rem',
            fontWeight: 600,
            color: typeColors[qType] || '#6b7280',
            textTransform: 'uppercase',
            letterSpacing: '0.05em',
          }}>
            {qTypeLabel}
          </span>
        );
      },
    },
    {
      field: 'isRequired',
      headerName: 'Required',
      width: 85,
      sortable: false,
      renderCell: (p: GridRenderCellParams) => {
        const isReq = (p.row as Question).isRequired === 'Yes';
        return (
          <Chip
            label={isReq ? 'Required' : 'Optional'}
            size="small"
            color={isReq ? 'error' : 'default'}
            variant={isReq ? 'filled' : 'outlined'}
            sx={{ fontSize: '0.65rem', height: '20px', fontWeight: 600 }}
          />
        );
      },
    },
    {
      field: 'status',
      headerName: 'Status',
      width: 150,
      sortable: false,
      renderCell: (p: GridRenderCellParams) => {
        const q = p.row as Question;
        const saving = (rowSaveStates[q.id] ?? 'idle') === 'saving';
        const isOn = q.status === 'Active';

        return (
          <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
            <Switch
              checked={isOn}
              onChange={() => void handleToggleStatus(q)}
              disabled={saving}
              size="small"
              sx={SWITCH_SX}
            />
            {saving ? (
              <CircularProgress size={14} sx={{ color: '#00D264' }} />
            ) : (
              <span style={{ fontSize: '0.75rem', fontWeight: 600, color: isOn ? '#004632' : '#6b7280' }}>
                {getStatusLabel(q.status)}
              </span>
            )}
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
            <IconButton
              className={styles.editActionBtn}
              size="small"
              onClick={() => openEditQuestionDrawer(question)}
              disabled={saving || isAdding}
              aria-label={`Update question: ${question.text}`}
              title="Update question"
            >
              <EditOutlinedIcon sx={{ fontSize: '1rem' }} />
            </IconButton>
            <IconButton
              size="small"
              onClick={() => void handleDeleteQuestion(question)}
              disabled={saving || isAdding}
              aria-label={`Delete question: ${question.text}`}
              title="Delete question"
              sx={{ color: '#dc2626', '&:hover': { backgroundColor: '#fee2e2' } }}
            >
              <DeleteOutlineIcon sx={{ fontSize: '1rem' }} />
            </IconButton>
          </div>
        );
      },
    },
  ];

  return (
    <div className={styles.Container} ref={containerRef}>
      <BuddyLoader
        visible={loaderVisible}
        variant="default"
        message={loaderMessage}
      />

      <div className={styles.pageHeading}>
        <div className={styles.pageHeadingLeft}>
          <div className={styles.pageHeadingIcon}>
            <HelpOutlineIcon sx={{ fontSize: '1.1rem' }} />
          </div>
          <div>
            <h1 className={styles.pageTitle}>Joinee Onboarding Questions</h1>
            <p className={styles.pageSubtitle}>Manage questions shown to new joinee employees during onboarding</p>
          </div>
        </div>
      </div>

      {!loadError && (
        <div className={styles.milestoneHeader}>
          <div className={styles.milestoneHeaderLeft}>
            <div className={styles.milestoneHeaderIcon}>
              <span>📋</span>
            </div>
            <div>
              <h2 className={styles.milestoneHeaderTitle}>Milestone Questions</h2>
              <p className={styles.milestoneHeaderSubtitle}>
                {activeQuestions} active of {totalQuestions} total questions
              </p>
            </div>
          </div>

          <div className={styles.milestoneTabsContainer}>
            {milestones.map((milestone) => {
              const stats = getMilestoneStats(milestone);
              const isSelected = selectedMilestone === milestone;
              return (
                <button
                  key={milestone}
                  className={`${styles.milestoneTab} ${isSelected ? styles.milestoneTabActive : ''}`}
                  onClick={() => {
                    setSelectedMilestone(milestone);
                    setSearchText('');
                  }}
                >
                  <div className={styles.milestoneTabLabel}>{milestone} Meeting</div>
                  <div className={styles.milestoneTabStats}>
                    <span className={styles.milestoneStatActive}>{stats.active}</span>
                    <span className={styles.milestoneStatDivider}>/</span>
                    <span className={styles.milestoneStatTotal}>{stats.total}</span>
                  </div>
                </button>
              );
            })}
          </div>
        </div>
      )}

      {loadError && (
        <div className={styles.errorBanner}>
          {loadError}
          <button className={styles.retryBtn} onClick={() => void loadQuestions()}>
            Retry
          </button>
        </div>
      )}

      {!loadError && (
        <Drawer
          anchor="right"
          open={isAddDrawerOpen}
          onClose={closeAddQuestionDrawer}
          ModalProps={{ disablePortal: true, keepMounted: true, container: containerRef.current }}
          sx={{
            ...DRAWER_OVER_TOPBAR_SX,
            "& .MuiDrawer-paper": {
              width: { xs: "100%", sm: "750px" },
              boxShadow: "-4px 0 28px rgba(0,0,0,0.18)",
              ...DRAWER_OVER_TOPBAR_SX["& .MuiDrawer-paper"],
            },
          }}>
          <div className={styles.addQuestionDrawer}>
            <div className={styles.addDrawerHeader}>
              <div>
                <h2 className={styles.addDrawerTitle}>
                  {editingQuestion ? 'Update Question' : 'Add New Question'}
                </h2>
                <p className={styles.addDrawerSub}>
                  {editingQuestion
                    ? 'Update the question while preserving its existing response mapping.'
                    : 'Create a feedback question from the master list form.'}
                </p>
              </div>
              <IconButton onClick={closeAddQuestionDrawer} disabled={isAdding} className={styles.addDrawerCloseBtn}>
                <CloseIcon sx={{ fontSize: '1rem' }} />
              </IconButton>
            </div>

            <div className={styles.addDrawerBody}>
              <div className={styles.addCardBadge}>{selectedMilestone} Meeting</div>

              <div className={styles.addCardField}>
                <label className={styles.addCardLabel}>
                  Milestone <span className={styles.required}>*</span>
                </label>
                <select
                  className={styles.selectInput}
                  value={selectedMilestone}
                  onChange={(e) => setSelectedMilestone(e.target.value)}
                  disabled={isAdding || Boolean(editingQuestion)}
                >
                  <option value="3rd">3rd Meeting</option>
                  <option value="6th">6th Meeting</option>
                </select>
              </div>

              <div className={styles.addCardField}>
                <label className={styles.addCardLabel}>
                  Question Type <span className={styles.required}>*</span>
                </label>
                <select
                  className={styles.selectInput}
                  value={selectedQuestionType}
                  onChange={(e) => setSelectedQuestionType(e.target.value)}
                  disabled={isAdding}
                >
                  <option value="yes_no">{QUESTION_TYPES.yes_no.label}</option>
                  <option value="agree_disagree">{QUESTION_TYPES.agree_disagree.label}</option>
                  <option value="rating">{QUESTION_TYPES.rating.label}</option>
                  <option value="remark">{QUESTION_TYPES.remark.label}</option>
                </select>
              </div>

              <div className={styles.addCardField}>
                <label className={styles.addCardLabel}>
                  Required <span className={styles.required}>*</span>
                </label>
                <select
                  className={styles.selectInput}
                  value={isRequired}
                  onChange={(e) => setIsRequired(e.target.value)}
                  disabled={isAdding}
                >
                  <option value="No">No</option>
                  <option value="Yes">Yes</option>
                </select>
              </div>

              <div className={styles.addCardField}>
                <label className={styles.addCardLabel}>Status</label>
                <div style={{ display: 'flex', alignItems: 'center', gap: '0.75rem' }}>
                  <span style={{ fontSize: '0.775rem', fontWeight: 600, color: '#6b7280' }}>Inactive</span>
                  <Switch
                    checked={isActive}
                    onChange={(e) => setIsActive(e.target.checked)}
                    disabled={isAdding}
                    size="small"
                    sx={SWITCH_SX}
                  />
                  <span style={{ fontSize: '0.775rem', fontWeight: 600, color: '#004632' }}>Active</span>
                </div>
              </div>

              <div className={styles.addCardField}>
                <label className={styles.addCardLabel}>
                  Question Text <span className={styles.required}>*</span>
                </label>
                <textarea
                  className={`${styles.textarea} ${duplicateWarning ? styles.textareaError : ''}`}
                  placeholder={`Enter the question for ${selectedMilestone} Meeting...`}
                  value={newQuestionText}
                  onChange={(e) => {
                    setNewQuestionText(e.target.value);
                    setDuplicateWarning('');
                  }}
                  disabled={isAdding}
                />
                {duplicateWarning && (
                  <div className={styles.duplicateWarning}>
                    ⚠️ {duplicateWarning}
                  </div>
                )}
              </div>
            </div>

            <div className={styles.addDrawerFooter}>
              <button className={styles.cancelBtn} onClick={closeAddQuestionDrawer} disabled={isAdding}>Cancel</button>
              <button
                className={styles.addBtn}
                onClick={() => void handleSaveQuestion()}
                disabled={!normalizeSingleSpaces(newQuestionText) || isAdding}>
                {isAdding
                  ? <CircularProgress size={16} />
                  : editingQuestion ? 'Update Question' : '+ Add Question'}
              </button>
            </div>
          </div>
        </Drawer>
      )}

      {!loadError && (
        <div className={styles.tableSection}>
          <div className={styles.tableHeader}>
            <div className={styles.tableHeaderLeft}>
              <h2 className={styles.tableTitle}>Questions — {selectedMilestone} Meeting</h2>
              <div className={styles.tableStats}>
                <span className={styles.tableStatActive}>{activeQuestions}</span>
                <span className={styles.tableStatDivider}>·</span>
                <span className={styles.tableStatTotal}>{totalQuestions} total</span>
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
              <button
                className={styles.openAddDrawerBtn}
                onClick={openAddQuestionDrawer}
                disabled={isLoading}
                type="button">
                + Add New Question
              </button>
              <div className={styles.searchWrap}>
                <input
                  type="text"
                  className={styles.searchInput}
                  placeholder="Search questions..."
                  value={searchText}
                  onChange={(e) => setSearchText(e.target.value)}
                />
                {searchText && (
                  <button
                    className={styles.searchClear}
                    onClick={() => setSearchText('')}
                    aria-label="Clear search"
                  >
                    <ClearIcon sx={{ fontSize: '1rem' }} />
                  </button>
                )}
              </div>
              {searchText && (
                <span className={styles.tableStatSearch}>
                  {visibleRows.length} match{visibleRows.length !== 1 ? 'es' : ''}
                </span>
              )}
              <div className={styles.statusLegend}>
                <RadioGroup
                  row
                  value={selectedStatus}
                  onChange={(e) => setSelectedStatus(e.target.value as 'All' | 'Active' | 'InActive')}
                  sx={{
                    display: 'flex',
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

          {isLoading ? (
            <div className={styles.loadingState}>
              <CircularProgress size={40} sx={{ color: BRAND.classicGreen }} />
              <p style={{ marginTop: '0.5rem', fontSize: '0.775rem', color: '#6b7280' }}>
                Loading questions...
              </p>
            </div>
          ) : visibleRows.length === 0 ? (
            <div className={styles.emptyState}>
              <p style={{ fontSize: '0.875rem', color: '#6b7280', margin: 0 }}>
                {searchText
                  ? 'No questions match your search. '
                  : `No questions added yet for ${selectedMilestone} Meeting. `}
                {searchText && (
                  <button
                    className={styles.clearSearchLink}
                    onClick={() => setSearchText('')}
                  >
                    Clear search
                  </button>
                )}
              </p>
            </div>
          ) : (
            <div className={styles.gridCard}>
              <DataGrid
                rows={visibleRows}
                columns={columns}
                disableRowSelectionOnClick
                hideFooterSelectedRowCount
                initialState={{ pagination: { paginationModel: { pageSize: 10 } } }}
                pageSizeOptions={[5, 10, 20, 50]}
                rowHeight={56}
                columnHeaderHeight={40}
                sx={{
                  border: 'none',
                  width: '100%',
                  height: '100%',
                  boxSizing: 'border-box',
                  willChange: 'transform',
                  '& .MuiDataGrid-columnHeaders': {
                    backgroundColor: BRAND.darkGreen,
                    color: '#fff',
                    fontSize: '0.7rem',
                    fontWeight: 700, 
                  },
                  '& .MuiDataGrid-columnHeaderTitle': {
                    fontWeight: 700,
                    textTransform: 'uppercase',
                    letterSpacing: '0.05em',
                    color: '#fff',
                  },
                  '& .MuiDataGrid-sortIcon': { color: '#fff' },
                  '& .MuiDataGrid-menuIconButton': { color: '#fff' },
                  '& .MuiDataGrid-columnSeparator': { color: 'rgba(255,255,255,0.15)' },
                  '& .MuiDataGrid-row:hover': { backgroundColor: '#F5F2ED' },
                  '& .MuiDataGrid-cell': {
                    fontSize: '0.775rem',
                    padding: '0 0.75rem',
                    borderColor: '#f3f4f6',
                    alignItems: 'center',
                  },
                  '& .MuiDataGrid-footerContainer': {
                    borderTop: '1px solid #e5e7eb',
                  },
                  '& .MuiTablePagination-root': { fontSize: '0.75rem' },
                  '& .MuiTablePagination-selectLabel, & .MuiTablePagination-displayedRows': {
                    fontSize: '0.725rem',
                  },
                }}
              />
            </div>
          )}
        </div>
      )}

      {/* <Snackbar
        open={Boolean(deleteTarget)}
        anchorOrigin={{ vertical: 'bottom', horizontal: 'right' }}
        onClose={() => setDeleteTarget(null)}
        message={
          <div style={{ maxWidth: 420 }}>
            <div style={{ fontWeight: 700, marginBottom: 4 }}>Remove this question?</div>
            <div style={{ fontSize: '0.8rem', lineHeight: 1.45 }}>
              It will no longer appear in the {deleteTarget?.milestone || selectedMilestone} Meeting questionnaire. Previously submitted responses will remain unchanged.
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
{/* delete question confirmation */}
      {renderDeleteConfirmation()}
    </div>
  );
};

export default JoineeQuestionnaire;