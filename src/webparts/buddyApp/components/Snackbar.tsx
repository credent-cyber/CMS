import * as React from 'react';
import { useState, useCallback } from 'react';
import { X, AlertCircle, CheckCircle } from 'lucide-react';
import styles from './Snackbar.module.scss';

export type SnackbarType = 'success' | 'error' | 'info' | 'warning';

const snackbarTypeMap: Record<SnackbarType, keyof typeof styles> = {
  success: 'snackbarSuccess',
  error: 'snackbarError',
  info: 'snackbarInfo',
  warning: 'snackbarWarning',
};

export interface SnackbarMessage {
  id: string;
  message: string;
  type: SnackbarType;
}

interface SnackbarContextType {
  showSnackbar: (message: string, type?: SnackbarType, duration?: number) => void;
}

const SnackbarContext = React.createContext<SnackbarContextType | undefined>(undefined);

export const useSnackbar = (): SnackbarContextType => {
  const context = React.useContext(SnackbarContext);
  if (!context) {
    throw new Error('useSnackbar must be used within SnackbarProvider');
  }
  return context;
};

export const SnackbarProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const [messages, setMessages] = useState<SnackbarMessage[]>([]);

  const removeSnackbar = useCallback((id: string) => {
    setMessages(prev => prev.filter(msg => msg.id !== id));
  }, []);
  
  const showSnackbar = useCallback((message: string, type: SnackbarType = 'info', duration: number = 3000) => {
    const id = Math.random().toString(36).substr(2, 9);
    const newMessage: SnackbarMessage = { id, message, type };

    setMessages(prev => [...prev, newMessage]);

    if (duration > 0) {
      setTimeout(() => {
        removeSnackbar(id);
      }, duration);
    }
  }, []);


  return (
    <SnackbarContext.Provider value={{ showSnackbar }}>
      {children}
      <div className={styles.snackbarContainer}>
        {messages.map(msg => (
          <div key={msg.id} className={`${styles.snackbar} ${styles[snackbarTypeMap[msg.type]]}`}>
            <div className={styles.snackbarContent}>
              {msg.type === 'success' && <CheckCircle size={20} className={styles.icon} />}
              {msg.type === 'error' && <AlertCircle size={20} className={styles.icon} />}
              <span className={styles.message}>{msg.message}</span>
            </div>
            <button
              className={styles.closeBtn}
              onClick={() => removeSnackbar(msg.id)}
              aria-label="Close snackbar"
            >
              <X size={18} />
            </button>
          </div>
        ))}
      </div>
    </SnackbarContext.Provider>
  );
};
