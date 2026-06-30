
import { motion } from 'motion/react';
import { AssessmentItem } from '../types';
import { cn } from '@/src/lib/utils';
import { Trash2 } from 'lucide-react';

interface Props {
  item: AssessmentItem;
  index: number;
  removeItem: (index: number) => void;
}

export const AssessmentSummaryItem = ({ item, index, removeItem }: Props) => {
  return (
    <motion.div 
      initial={{ opacity: 0, y: 10 }}
      animate={{ opacity: 1, y: 0 }}
      key={index}
      className="p-3 bg-slate-50 rounded-xl border border-slate-100 group relative"
    >
      <button 
        onClick={() => removeItem(index)}
        className="absolute -top-2 -right-2 w-6 h-6 bg-white border border-slate-200 rounded-full flex items-center justify-center text-red-500 shadow-sm opacity-0 group-hover:opacity-100 transition-opacity"
      >
        <Trash2 className="w-3 h-3" />
      </button>
      <div className="flex justify-between items-start mb-1">
        <p className="font-bold text-sm text-slate-900 line-clamp-1">{item.material.name}</p>
        <span className={cn(
          "text-[10px] font-bold px-1.5 py-0.5 rounded uppercase",
          item.status === 'damaged' ? "bg-red-100 text-red-700" : "bg-green-100 text-green-700"
        )}>
          {item.status === 'damaged' ? 'ชำรุด' : 'นำกลับมาใช้ใหม่'}
        </span>
      </div>
      <div className="flex justify-between items-end">
        <p className="text-xs text-slate-500">
          {item.quantity} {item.material.unit} x ฿{(item.status === 'damaged' ? item.material.priceDamaged : item.material.priceReusable).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
        </p>
        <p className="font-bold text-blue-600">฿{item.totalPrice.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</p>
      </div>
    </motion.div>
  );
};
