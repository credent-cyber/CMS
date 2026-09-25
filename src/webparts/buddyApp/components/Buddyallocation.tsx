/* eslint-disable @typescript-eslint/explicit-function-return-type*/
/* eslint-disable react/self-closing-comp */
import * as React from 'react';
import { useRef, useState } from 'react';
import { IconButton, Tooltip, Drawer } from '@mui/material';
import DeleteOutlineIcon from '@mui/icons-material/DeleteOutline';
import CloseIcon from '@mui/icons-material/Close';
import CalendarTodayIcon from '@mui/icons-material/CalendarToday';
import PersonAddAlt1Icon from '@mui/icons-material/PersonAddAlt1';
import styles from './buddyAllocation.module.scss';

const BRAND = {
  darkGreen:    "#004632",   
  classicGreen: "#00D264",   
  lightGreen:   "#8CFF8C",   
  white:        "#FFFFFF",
  lightStone:   "#F5F2ED",   
  darkStone:    "#78736E",   
} as const;


interface BuddyAllocation {
  id: string;
  name: string;
  doj: string;
  meeting1: string;
  meeting2: string;
  meeting3: string;
  meeting4: string;
  meeting5: string;
  meeting6: string;
}


export const BuddyAllocation: React.FC = () => {
  const containerRef = useRef<HTMLDivElement | null>(null);

  const [drawerOpen, setDrawerOpen] = useState(false);
  const [buddyName,   setBuddyName]   = useState('');
  const [department,  setDepartment]  = useState('');
  const [njName,      setNjName]      = useState('');
  const [doj,         setDoj]         = useState('');
  const [globalId,    setGlobalId]    = useState('');
  const [allocations, setAllocations] = useState<BuddyAllocation[]>([]);

  const handleCloseDrawer = () => {
    setDrawerOpen(false);
    setBuddyName('');
    setDepartment('');
    setNjName('');
    setDoj('');
    setGlobalId('');
  };

  const handleSave = () => {
    handleCloseDrawer();
  };

  const handleDelete = (id: string) => {
    setAllocations(prev => prev.filter(item => item.id !== id));
  };


  return (
    <div className={styles.buddyAllocationContainer} ref={containerRef}>

      <main className={styles.main}>

        <div className={styles.toolbar}>
          <button className={styles.assignBtn} onClick={() => setDrawerOpen(true)}>
            <PersonAddAlt1Icon sx={{ fontSize: '1.1rem' }} />
            <span>Assign Buddy</span>
          </button>
        </div>

        <div className={styles.tableCard}>
          <div className={styles.tableWrapper}>
            <table className={styles.table}>
              <thead>
                <tr>
                  <th>NAME</th>
                  <th>DOJ</th>
                  <th>1ST MEETING</th>
                  <th>2ND MEETING</th>
                  <th>3RD MEETING</th>
                  <th>4TH MEETING</th>
                  <th>5TH MEETING</th>
                  <th>6TH MEETING</th>
                  <th>ACTION</th>
                </tr>
              </thead>
              <tbody>
                {allocations.length === 0 ? (
                  <tr>
                    <td colSpan={9} className={styles.emptyRow}>
                      No allocations yet. Click &ldquo;Assign Buddy&rdquo; to add one.
                    </td>
                  </tr>
                ) : (
                  allocations.map((allocation) => (
                    <tr key={allocation.id}>
                      <td>{allocation.name}</td>
                      <td>{allocation.doj}</td>
                      <td>{allocation.meeting1}</td>
                      <td>{allocation.meeting2}</td>
                      <td>{allocation.meeting3}</td>
                      <td>{allocation.meeting4}</td>
                      <td>{allocation.meeting5}</td>
                      <td>{allocation.meeting6}</td>
                      <td>
                        <Tooltip title="Delete">
                          <IconButton
                            size="small"
                            onClick={() => handleDelete(allocation.id)}
                            className={styles.deleteBtn}
                          >
                            <DeleteOutlineIcon sx={{ fontSize: '1.1rem' }} />
                          </IconButton>
                        </Tooltip>
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        </div>
      </main>

      <Drawer
        anchor="right"
        open={drawerOpen}
        onClose={handleCloseDrawer}
        ModalProps={{
          disablePortal: true,
          keepMounted: true,
          container: containerRef.current,
        }}
        sx={{
          '& .MuiDrawer-paper': {
            width: { xs: '100%', sm: '500px', md: '550px' },
            boxShadow: '-4px 0 20px rgba(0, 0, 0, 0.15)',
            // Light Stone background — on-brand drawer surface
            background: BRAND.lightStone,
            fontFamily: "'Inter', sans-serif",
          },
        }}
      >
        <div className={styles.drawer}>

          <div className={styles.drawerHeader}>
            <div>
              <h2 className={styles.drawerTitle}>Assign Buddy</h2>
              <p className={styles.drawerSubtitle}>Fill in the details to assign a buddy</p>
            </div>
            <IconButton onClick={handleCloseDrawer} className={styles.closeBtn}>
              <CloseIcon />
            </IconButton>
          </div>

          <div className={styles.drawerBody}>

            <div className={styles.formGroup}>
              <label className={styles.label}>
                Buddy Name <span className={styles.required}>*</span>
              </label>
              <select
                value={buddyName}
                onChange={(e) => setBuddyName(e.target.value)}
                className={styles.select}
              >
                <option value="">Search Buddy</option>
              </select>
            </div>

            <div className={styles.formGroup}>
              <label className={styles.label}>Department</label>
              <input
                type="text"
                value={department}
                onChange={(e) => setDepartment(e.target.value)}
                className={styles.input}
                placeholder="Enter department"
              />
            </div>

            <div className={styles.formGroup}>
              <label className={styles.label}>
                NJ Name <span className={styles.required}>*</span>
              </label>
              <select
                value={njName}
                onChange={(e) => setNjName(e.target.value)}
                className={styles.select}
              >
                <option value="">Search New Joiner</option>
              </select>
            </div>

            <div className={styles.formGroup}>
              <label className={styles.label}>
                DOJ <span className={styles.required}>*</span>
              </label>
              <div className={styles.dateInputWrapper}>
                <input
                  type="text"
                  value={doj}
                  onChange={(e) => setDoj(e.target.value)}
                  className={styles.input}
                  placeholder="Monday, December 31, 2025"
                />
                <CalendarTodayIcon className={styles.calendarIcon} />
              </div>
            </div>

            <div className={styles.formGroup}>
              <label className={styles.label}>Global ID</label>
              <input
                type="text"
                value={globalId}
                onChange={(e) => setGlobalId(e.target.value)}
                className={styles.input}
                placeholder="Enter Global ID"
              />
            </div>

          </div>

          <div className={styles.drawerFooter}>
            <button className={styles.cancelBtn} onClick={handleCloseDrawer}>
              Cancel
            </button>
            <button
              className={styles.saveBtn}
              onClick={handleSave}
              disabled={!buddyName || !njName || !doj}
            >
              Save
            </button>
          </div>

        </div>
      </Drawer>
    </div>
  );
};

export default BuddyAllocation;