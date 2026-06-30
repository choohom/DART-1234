
import { Document, Paragraph, Table, TableCell, TableRow, WidthType, AlignmentType } from 'docx';
import { AssessmentItem } from '../types';

export const groupItems = (itemsList: AssessmentItem[]) => {
    const grouped: { [key: string]: AssessmentItem } = {};
    itemsList.forEach(item => {
        const key = item.material.name;
        if (grouped[key]) {
            grouped[key] = {
                ...grouped[key],
                quantity: grouped[key].quantity + item.quantity,
                totalPrice: grouped[key].totalPrice + item.totalPrice
            };
        } else {
            grouped[key] = { ...item };
        }
    });
    return Object.values(grouped);
};

export const fetchImage = async (url: string) => {
    try {
        const response = await fetch(url);
        const blob = await response.blob();
        return await blob.arrayBuffer();
    } catch (error) {
        console.error("Error fetching logo:", error);
        return null;
    }
};
